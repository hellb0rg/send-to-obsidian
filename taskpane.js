'use strict';

/* ═══════════════════════════════════════════════════════════════════════════
   Outlook to Obsidian — Office Add-in task pane
   ═══════════════════════════════════════════════════════════════════════════
   Uses Office.js (Office.context.mailbox.item) to read meeting/email data.

   This version separates Outlook data extraction from note presentation:
     templates/event-template.md
     templates/mail-template.md

   Template placeholders use the form {{variable_name}}.

   RSVP status is fetched via EWS (makeEwsRequestAsync) when possible,
   preserving the behavior of the upstream Virtual-Jones implementation.
   If EWS is unavailable, the add-in falls back to Office.js attendee data.
   ═══════════════════════════════════════════════════════════════════════════ */


/* ── Defaults ────────────────────────────────────────────────────────────── */

const DEFAULT_SETTINGS = {
  vaultName:       'Notes',
  folderPath:      '02. Meeting Notes',
  emailFolderPath: '04. Email Notes',
  internalDomain:  'example.com',
  copyToClipboard: true,
  includeExternal: false,
};

let settings = { ...DEFAULT_SETTINGS };
let extractedData = null;


/* ═══════════════════════════════════════════════════════════════════════════
   BOOTSTRAP
   ═══════════════════════════════════════════════════════════════════════════ */

Office.onReady(() => {
  loadSettings();
  bindUI();
  applyContextLabels();
});


/* Returns 'message' or 'meeting' based on the current Outlook item. */
function getCurrentItemType() {
  try {
    const t = Office.context.mailbox.item?.itemType;
    return t === Office.MailboxEnums.ItemType.Message ? 'message' : 'meeting';
  } catch (_) {
    return 'meeting';
  }
}


/* Update button label and header to match the current context. */
function applyContextLabels() {
  const isMessage = getCurrentItemType() === 'message';

  const btn = document.getElementById('extractBtn');
  if (btn) {
    btn.innerHTML = isMessage
      ? '<span>📧</span><span>Extract Email Details</span>'
      : '<span>📅</span><span>Extract Meeting Details</span>';
  }

  const headerIcon = document.getElementById('headerIcon');
  if (headerIcon) {
    headerIcon.textContent = isMessage ? '📧' : '📅';
  }

  const headerTitle = document.getElementById('headerTitle');
  if (headerTitle) {
    headerTitle.textContent = isMessage
      ? 'Send Email to Obsidian'
      : 'Send Meeting to Obsidian';
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   UI BINDING
   ═══════════════════════════════════════════════════════════════════════════ */

function bindUI() {
  // Main panel
  document.getElementById('extractBtn')
    .addEventListener('click', onExtract);

  document.getElementById('openObsidianBtn')
    .addEventListener('click', onOpenObsidian);

  document.getElementById('copyBtn')
    .addEventListener('click', onCopy);

  document.getElementById('settingsBtn')
    .addEventListener('click', () => showPanel('settings'));

  // Settings panel
  document.getElementById('backBtn')
    .addEventListener('click', () => showPanel('main'));

  document.getElementById('saveSettingsBtn')
    .addEventListener('click', onSaveSettings);

  // Preset buttons
  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target;
      const value  = btn.dataset.value;

      if (target && value) {
        document.getElementById(target).value = value;

        if (target === 'vaultName' || target === 'folderPath') {
          updateFolderPreview();
        }
      }
    });
  });

  // Live preview in settings
  document.getElementById('vaultName')
    .addEventListener('input', updateFolderPreview);

  document.getElementById('folderPath')
    .addEventListener('input', updateFolderPreview);
}


/* ═══════════════════════════════════════════════════════════════════════════
   PANEL SWITCHING
   ═══════════════════════════════════════════════════════════════════════════ */

function showPanel(name) {
  const isMain = name === 'main';

  document.getElementById('mainPanel').style.display =
    isMain ? 'flex' : 'none';

  document.getElementById('settingsPanel').style.display =
    isMain ? 'none' : 'flex';

  if (!isMain) {
    populateSettingsForm();
    updateFolderPreview();
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   SETTINGS
   ═══════════════════════════════════════════════════════════════════════════ */

function loadSettings() {
  try {
    const rs = Office.context.roamingSettings;

    for (const key of Object.keys(DEFAULT_SETTINGS)) {
      const value = rs.get(key);

      if (value !== undefined && value !== null) {
        settings[key] = value;
      }
    }
  } catch (_) {
    // roamingSettings unavailable — use defaults.
  }

  updateConfigDisplay();
}


function onSaveSettings() {
  const vaultName =
    document.getElementById('vaultName').value.trim();

  const folderPath =
    document.getElementById('folderPath').value.trim();

  const emailFolderPath =
    document.getElementById('emailFolderPath').value.trim();

  const internalDomain =
    document.getElementById('internalDomain').value.trim();

  if (!vaultName) {
    showStatus('Please enter a vault name.', 'error');
    return;
  }

  if (!folderPath) {
    showStatus('Please enter a meeting notes folder.', 'error');
    return;
  }

  if (!emailFolderPath) {
    showStatus('Please enter an email notes folder.', 'error');
    return;
  }

  settings.vaultName       = vaultName;
  settings.folderPath      = folderPath;
  settings.emailFolderPath = emailFolderPath;
  settings.internalDomain  =
    internalDomain || DEFAULT_SETTINGS.internalDomain;

  settings.copyToClipboard =
    document.getElementById('copyToClipboard').checked;

  settings.includeExternal =
    document.getElementById('includeExternal').checked;

  try {
    const rs = Office.context.roamingSettings;

    for (const [key, value] of Object.entries(settings)) {
      rs.set(key, value);
    }

    rs.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        updateConfigDisplay();
        showPanel('main');
        showStatus('Settings saved.', 'success');
      } else {
        showStatus(
          'Could not save settings: ' +
            (result.error?.message || 'unknown error'),
          'error'
        );
      }
    });
  } catch (e) {
    showStatus('Error saving settings: ' + e.message, 'error');
  }
}


function populateSettingsForm() {
  document.getElementById('vaultName').value =
    settings.vaultName;

  document.getElementById('folderPath').value =
    settings.folderPath;

  document.getElementById('emailFolderPath').value =
    settings.emailFolderPath;

  document.getElementById('internalDomain').value =
    settings.internalDomain;

  document.getElementById('copyToClipboard').checked =
    settings.copyToClipboard;

  document.getElementById('includeExternal').checked =
    settings.includeExternal;
}


function updateConfigDisplay() {
  const isMessage = getCurrentItemType() === 'message';

  const activePath = isMessage
    ? settings.emailFolderPath
    : settings.folderPath;

  const ok = !!(settings.vaultName && activePath);

  document.getElementById('configWarning').style.display =
    ok ? 'none' : 'block';

  document.getElementById('vaultDisplay').textContent =
    settings.vaultName || 'Not set';

  document.getElementById('pathDisplay').textContent =
    activePath || 'Not set';

  document.getElementById('extractBtn').disabled = !ok;
}


function updateFolderPreview() {
  const vault =
    document.getElementById('vaultName').value || 'YourVault';

  const folder =
    document.getElementById('folderPath').value || 'Meeting Notes';

  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');

  document.getElementById('folderPreview').textContent =
    `📁 ${vault}/\n` +
    `  📁 ${folder}/\n` +
    `    📁 ${y}/\n` +
    `      📁 ${m}/\n` +
    `        📄 ${y}-${m}-15 Example Meeting.md`;
}


/* ═══════════════════════════════════════════════════════════════════════════
   TEMPLATE SYSTEM
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Loads a Markdown template from ./templates/.
 *
 * GitHub Pages example:
 *   /outlook-to-obsidian/templates/event-template.md
 *   /outlook-to-obsidian/templates/mail-template.md
 */
async function loadNoteTemplate(filename) {
  const response = await fetch(`./templates/${filename}`, {
    cache: 'no-store',
  });

  if (!response.ok) {
    throw new Error(
      `Could not load note template "${filename}" ` +
      `(${response.status} ${response.statusText})`
    );
  }

  return await response.text();
}


/**
 * Replaces {{variable_name}} placeholders.
 *
 * Unknown variables are replaced with an empty string and logged so that
 * a typo in the Markdown template does not appear in the final note.
 */
function renderNoteTemplate(template, values) {
  return template.replace(
    /\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g,
    (_match, key) => {
      if (!Object.prototype.hasOwnProperty.call(values, key)) {
        console.warn(`Unknown template variable: ${key}`);
        return '';
      }

      return values[key] ?? '';
    }
  );
}


/**
 * Quote a scalar safely for normal YAML frontmatter.
 */
function yamlQuote(value) {
  return `"${String(value ?? '')
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\r?\n/g, '\\n')}"`;
}


/**
 * Format an array as a YAML list property.
 *
 * Example:
 * attendees:
 *   - "Anna Andersson"
 *   - "Erik Eriksson"
 */
function yamlListProperty(propertyName, values) {
  const unique = Array.from(
    new Set((values || []).filter(Boolean))
  );

  if (unique.length === 0) {
    return `${propertyName}: []`;
  }

  return [
    `${propertyName}:`,
    ...unique.map(value => `  - ${yamlQuote(value)}`),
  ].join('\n');
}


/* ═══════════════════════════════════════════════════════════════════════════
   EXTRACT BUTTON HANDLER
   ═══════════════════════════════════════════════════════════════════════════ */

async function onExtract() {
  const btn = document.getElementById('extractBtn');

  btn.disabled = true;

  document.getElementById('details').style.display = 'none';
  document.getElementById('actions').style.display = 'none';

  const isMessage = getCurrentItemType() === 'message';

  showStatus(
    isMessage
      ? 'Extracting email…'
      : 'Extracting meeting details…',
    'info'
  );

  try {
    extractedData = isMessage
      ? await extractEmailDetails()
      : await extractMeetingDetails();

    // Apply [EXTERNAL] prefix when configured.
    //
    // The note H1 is updated after rendering. The generated filename also uses
    // extractedData.title later in buildObsidianUri(), so it receives the same
    // prefix automatically.
    if (
      settings.includeExternal &&
      extractedData.hasExternalAttendees
    ) {
      extractedData.title =
        '[EXTERNAL] ' + extractedData.title;

      extractedData.note = extractedData.note.replace(
        /^(# ).+$/m,
        `$1${extractedData.title}`
      );
    }

    renderDetails(extractedData);

    showStatus(
      isMessage
        ? 'Email extracted successfully.'
        : 'Meeting details extracted successfully.',
      'success'
    );

    if (settings.copyToClipboard) {
      await copyText(extractedData.note);
    }

    document.getElementById('details').style.display = 'block';
    document.getElementById('actions').style.display = 'flex';
  } catch (err) {
    showStatus('Error: ' + err.message, 'error');
  } finally {
    const activePath = isMessage
      ? settings.emailFolderPath
      : settings.folderPath;

    btn.disabled = !(settings.vaultName && activePath);
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   DETAILS PREVIEW
   ═══════════════════════════════════════════════════════════════════════════ */

function renderDetails(data) {
  const truncate = (value, maxLength) => {
    const s = String(value || '');

    return s.length > maxLength
      ? s.substring(0, maxLength) + '…'
      : s;
  };

  const isEmail = data.kind === 'email';
  const rows = [];

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Subject' : 'Title'}</span>
      <span class="detail-value">${esc(truncate(data.title, 60))}</span>
    </div>
  `);

  if (isEmail) {
    rows.push(`
      <div class="detail-row">
        <span class="detail-label">From</span>
        <span class="detail-value">${esc(truncate(data.fromName || '', 50))}</span>
      </div>
    `);
  }

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Sent' : 'Date/Time'}</span>
      <span class="detail-value">${esc(data.time || 'Not specified')}</span>
    </div>
  `);

  if (!isEmail && data.location) {
    rows.push(`
      <div class="detail-row">
        <span class="detail-label">Location</span>
        <span class="detail-value">${esc(truncate(data.location, 50))}</span>
      </div>
    `);
  }

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Recipients' : 'Attendees'}</span>
      <span class="detail-value">${data.totalAttendees}</span>
    </div>
  `);

  if (data.hasExternalAttendees) {
    rows.push(`
      <div class="detail-row">
        <span class="detail-label">External</span>
        <span class="detail-value external">Yes</span>
      </div>
    `);
  }

  document.getElementById('details').innerHTML =
    rows.join('');
}


/* ═══════════════════════════════════════════════════════════════════════════
   OPEN IN OBSIDIAN
   ═══════════════════════════════════════════════════════════════════════════ */

function onOpenObsidian() {
  if (!extractedData) {
    return;
  }

  const uri = buildObsidianUri(extractedData);
  const link = document.getElementById('obsidianLink');

  link.href = uri;

  // Clicking a hidden anchor is reliable across Outlook Desktop/Web.
  link.click();
}


function buildObsidianUri(data) {
  const basePath = data.kind === 'email'
    ? settings.emailFolderPath
    : settings.folderPath;

  const folder =
    `${basePath}/${data.meetingDate.year}/${data.meetingDate.month}`;

  const safeTitle =
    data.title.replace(/[<>:"/\\|?*]/g, '-');

  const file =
    `${folder}/${data.meetingDate.full} ${safeTitle}`;

  return (
    `obsidian://new` +
    `?vault=${encodeURIComponent(settings.vaultName)}` +
    `&file=${encodeURIComponent(file)}` +
    `&content=${encodeURIComponent(data.note)}`
  );
}


/* ═══════════════════════════════════════════════════════════════════════════
   CLIPBOARD
   ═══════════════════════════════════════════════════════════════════════════ */

async function onCopy() {
  if (!extractedData) {
    return;
  }

  await copyText(extractedData.note);
  showStatus('Note copied to clipboard.', 'success');
}


async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch (_) {
    // Fallback for older WebView2 builds.
    const ta = document.createElement('textarea');

    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity  = '0';

    document.body.appendChild(ta);

    ta.focus();
    ta.select();

    document.execCommand('copy');

    document.body.removeChild(ta);
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   UI HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */

function showStatus(message, type) {
  const el = document.getElementById('status');

  el.textContent = message;
  el.className = `status ${type}`;
  el.style.display = 'block';
}


function esc(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


/* ═══════════════════════════════════════════════════════════════════════════
   CORE EXTRACTION — MEETING
   ═══════════════════════════════════════════════════════════════════════════ */

async function extractMeetingDetails() {
  const item = Office.context.mailbox.item;

  // Read mode has a persisted item ID; compose mode normally does not.
  const isReadMode = !!item.itemId;


  /* ── Subject ──────────────────────────────────────────────────────────── */

  const title =
    (await getItemProperty(item.subject)) ||
    'Untitled Meeting';


  /* ── Start / end ──────────────────────────────────────────────────────── */

  const startDate =
    await getItemProperty(item.start);

  const endDate =
    await getItemProperty(item.end);

  let time = '';
  let startTime = '';
  let endTime = '';
  let meetingDate = null;

  if (
    startDate instanceof Date &&
    !isNaN(startDate)
  ) {
    const y =
      startDate.getFullYear();

    const mo =
      String(startDate.getMonth() + 1).padStart(2, '0');

    const d =
      String(startDate.getDate()).padStart(2, '0');

    meetingDate = {
      year: String(y),
      month: mo,
      day: d,
      full: `${y}-${mo}-${d}`,
    };

    const timeFormat = {
      hour: '2-digit',
      minute: '2-digit',
    };

    startTime =
      startDate.toLocaleTimeString([], timeFormat);

    endTime =
      endDate instanceof Date && !isNaN(endDate)
        ? endDate.toLocaleTimeString([], timeFormat)
        : '';

    time =
      `${startDate.toLocaleDateString()} ${startTime}` +
      (endTime ? ` – ${endTime}` : '');
  }

  if (!meetingDate) {
    const today = new Date();

    meetingDate = {
      year: String(today.getFullYear()),
      month:
        String(today.getMonth() + 1).padStart(2, '0'),
      day:
        String(today.getDate()).padStart(2, '0'),
      full:
        `${today.getFullYear()}-` +
        `${String(today.getMonth() + 1).padStart(2, '0')}-` +
        `${String(today.getDate()).padStart(2, '0')}`,
    };
  }


  /* ── Location ─────────────────────────────────────────────────────────── */

  const location =
    (await getItemProperty(item.location)) || '';


  /* ── Organizer ────────────────────────────────────────────────────────── */

  let organizer = '';
  let organizerEmail = '';

  if (isReadMode && item.organizer) {
    organizer =
      item.organizer.displayName || '';

    organizerEmail =
      (item.organizer.emailAddress || '')
        .toLowerCase();
  }


  /* ── Body / description ───────────────────────────────────────────────── */

  const body =
    await getBodyMarkdown(item);


  /* ── Attendees with RSVP ──────────────────────────────────────────────── */

  let attendeesByStatus = {
    'Accepted':    new Map(),
    'Tentative':   new Map(),
    'Declined':    new Map(),
    'No Response': new Map(),
  };

  let ewsErrorMessage = '';

  if (isReadMode && item.itemId) {
    try {
      attendeesByStatus =
        await getAttendeesViaEws(item.itemId);
    } catch (ewsErr) {
      console.warn(
        'EWS attendee fetch failed, using Office.js fallback:',
        ewsErr.message
      );

      ewsErrorMessage =
        ewsErr.message || 'unknown';

      attendeesByStatus =
        await getAttendeesViaOfficeJs(item);
    }
  } else {
    attendeesByStatus =
      await getAttendeesViaOfficeJs(item);
  }


  /* ── External attendee detection ──────────────────────────────────────── */

  const internalSuffix =
    '@' +
    (
      settings.internalDomain ||
      DEFAULT_SETTINGS.internalDomain
    )
      .toLowerCase()
      .replace(/^@/, '');

  let hasExternalAttendees =
    organizerEmail
      ? !organizerEmail.endsWith(internalSuffix)
      : false;

  if (!hasExternalAttendees) {
    for (
      const attendees
      of Object.values(attendeesByStatus)
    ) {
      for (const email of attendees.values()) {
        if (
          email &&
          !email.endsWith(internalSuffix)
        ) {
          hasExternalAttendees = true;
          break;
        }
      }

      if (hasExternalAttendees) {
        break;
      }
    }
  }


  /* ── Format attendee information ──────────────────────────────────────── */

  const attendeesFormatted = [];
  const allAttendeeNames = [];

  let totalAttendees =
    organizer ? 1 : 0;

  if (organizer) {
    attendeesFormatted.push(
      `**Organizer:** ${organizer}`
    );

    allAttendeeNames.push(organizer);
  }

  for (
    const status
    of ['Accepted', 'Tentative', 'No Response', 'Declined']
  ) {
    const list =
      Array.from(
        attendeesByStatus[status].keys()
      ).sort();

    if (list.length > 0) {
      totalAttendees += list.length;

      attendeesFormatted.push(
        `**${status} (${list.length}):** ` +
        list.join(', ')
      );

      allAttendeeNames.push(...list);
    }
  }

  const attendeesStr =
    attendeesFormatted.join('\n\n') ||
    'No attendees found';

  const attendeesFrontmatter =
    yamlListProperty(
      'attendees',
      allAttendeeNames
    );


  /* ── Render event-template.md ─────────────────────────────────────────── */

  const template =
    await loadNoteTemplate('event-template.md');

  const note =
    renderNoteTemplate(template, {
      // General
      title,

      // Date
      date: meetingDate.full,
      year: meetingDate.year,
      month: meetingDate.month,
      day: meetingDate.day,

      // Time
      time: time || 'Not specified',
      start_time: startTime,
      end_time: endTime,
      start_time_yaml: yamlQuote(startTime),
      end_time_yaml: yamlQuote(endTime),

      // Location
      location,
      location_yaml: yamlQuote(location),

      // Organizer
      organizer,
      organizer_email: organizerEmail,
      organizer_yaml: yamlQuote(organizer),
      organizer_email_yaml: yamlQuote(organizerEmail),

      // Attendees
      attendees: attendeesStr,
      attendees_frontmatter: attendeesFrontmatter,
      attendee_count: String(totalAttendees),

      // RSVP lists
      accepted_attendees:
        Array.from(
          attendeesByStatus['Accepted'].keys()
        ).sort().join(', '),

      tentative_attendees:
        Array.from(
          attendeesByStatus['Tentative'].keys()
        ).sort().join(', '),

      no_response_attendees:
        Array.from(
          attendeesByStatus['No Response'].keys()
        ).sort().join(', '),

      declined_attendees:
        Array.from(
          attendeesByStatus['Declined'].keys()
        ).sort().join(', '),

      // Meeting classification
      external:
        String(hasExternalAttendees),

      external_line:
        hasExternalAttendees
          ? '**External Meeting:** Yes'
          : '',

      // Outlook meeting description / agenda
      body:
        body.trim() || '',
    });


  /* ── Return extracted meeting data ────────────────────────────────────── */

  const safeTitle =
    title.replace(/[<>:"/\\|?*]/g, '-');

  return {
    kind: 'meeting',
    title,
    time,
    startTime,
    endTime,
    location,
    organizer,
    organizerEmail,
    totalAttendees,
    hasExternalAttendees,
    meetingDate,
    safeTitle,
    note,
    ewsErrorMessage,
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   CORE EXTRACTION — EMAIL
   ═══════════════════════════════════════════════════════════════════════════ */

async function extractEmailDetails() {
  const item =
    Office.context.mailbox.item;


  /* ── Subject ──────────────────────────────────────────────────────────── */

  const title =
    (await getItemProperty(item.subject)) ||
    'No Subject';


  /* ── Sent date ────────────────────────────────────────────────────────── */

  const sentDate =
    item.dateTimeCreated instanceof Date &&
    !isNaN(item.dateTimeCreated)
      ? item.dateTimeCreated
      : (
          item.dateTimeModified instanceof Date &&
          !isNaN(item.dateTimeModified)
            ? item.dateTimeModified
            : new Date()
        );

  const y =
    sentDate.getFullYear();

  const mo =
    String(sentDate.getMonth() + 1)
      .padStart(2, '0');

  const d =
    String(sentDate.getDate())
      .padStart(2, '0');

  const meetingDate = {
    year: String(y),
    month: mo,
    day: d,
    full: `${y}-${mo}-${d}`,
  };

  const sentTime =
    sentDate.toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
    });

  const time =
    `${sentDate.toLocaleDateString()} ${sentTime}`;


  /* ── From / sender ────────────────────────────────────────────────────── */

  const fromObj =
    item.from || item.sender || {};

  const fromName =
    fromObj.displayName ||
    fromObj.emailAddress ||
    '';

  const fromEmail =
    (fromObj.emailAddress || '')
      .toLowerCase();


  /* ── To / Cc ──────────────────────────────────────────────────────────── */

  const toList =
    Array.isArray(item.to)
      ? item.to
      : [];

  const ccList =
    Array.isArray(item.cc)
      ? item.cc
      : [];

  function formatPerson(person) {
    const name =
      person.displayName ||
      person.emailAddress ||
      '';

    const email =
      person.emailAddress || '';

    return (
      email &&
      name !== email
        ? `${name} <${email}>`
        : name
    );
  }

  const toFormatted =
    toList.map(formatPerson).filter(Boolean);

  const ccFormatted =
    ccList.map(formatPerson).filter(Boolean);


  /* ── Body ─────────────────────────────────────────────────────────────── */

  const body =
    await getBodyMarkdown(item);


  /* ── External detection ───────────────────────────────────────────────── */

  const internalSuffix =
    '@' +
    (
      settings.internalDomain ||
      DEFAULT_SETTINGS.internalDomain
    )
      .toLowerCase()
      .replace(/^@/, '');

  let hasExternalAttendees =
    !!(
      fromEmail &&
      !fromEmail.endsWith(internalSuffix)
    );

  if (!hasExternalAttendees) {
    for (
      const person
      of [...toList, ...ccList]
    ) {
      const email =
        (person.emailAddress || '')
          .toLowerCase();

      if (
        email &&
        !email.endsWith(internalSuffix)
      ) {
        hasExternalAttendees = true;
        break;
      }
    }
  }

  const totalAttendees =
    toList.length + ccList.length;


  /* ── Sender display value ─────────────────────────────────────────────── */

  const fromLine =
    fromName
      ? (
          fromEmail &&
          fromName !== fromEmail
            ? `${fromName} <${fromEmail}>`
            : fromName
        )
      : 'Unknown sender';


  /* ── Render mail-template.md ──────────────────────────────────────────── */

  const template =
    await loadNoteTemplate('mail-template.md');

  const note =
    renderNoteTemplate(template, {
      // General
      title,

      // Date/time
      date: meetingDate.full,
      year: meetingDate.year,
      month: meetingDate.month,
      day: meetingDate.day,
      time,
      sent_time: sentTime,

      // Sender
      from: fromLine,
      from_name: fromName,
      from_email: fromEmail,
      from_yaml: yamlQuote(fromLine),
      from_name_yaml: yamlQuote(fromName),
      from_email_yaml: yamlQuote(fromEmail),

      // Recipients
      to: toFormatted.join(', '),
      cc: ccFormatted.join(', '),

      to_line:
        toFormatted.length
          ? `**To (${toFormatted.length}):** ` +
            toFormatted.join(', ')
          : '',

      cc_line:
        ccFormatted.length
          ? `**Cc (${ccFormatted.length}):** ` +
            ccFormatted.join(', ')
          : '',

      to_frontmatter:
        yamlListProperty('to', toFormatted),

      cc_frontmatter:
        yamlListProperty('cc', ccFormatted),

      recipient_count:
        String(totalAttendees),

      // Classification
      external:
        String(hasExternalAttendees),

      external_line:
        hasExternalAttendees
          ? '**External Email:** Yes'
          : '',

      // Body
      body:
        body.trim() || '',
    });


  /* ── Return extracted email data ──────────────────────────────────────── */

  const safeTitle =
    title.replace(/[<>:"/\\|?*]/g, '-');

  return {
    kind: 'email',
    title,
    time,
    fromName: fromLine,
    fromEmail,
    location: '',
    totalAttendees,
    hasExternalAttendees,
    meetingDate,
    safeTitle,
    note,
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   OFFICE.JS HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Reads a property that is either a direct value (read mode) or exposes
 * getAsync() (compose mode), and returns a Promise for the resolved value.
 */
function getItemProperty(propOrValue) {
  return new Promise(resolve => {
    if (
      propOrValue === undefined ||
      propOrValue === null
    ) {
      resolve('');
      return;
    }

    // Direct value: string, Date, number, boolean.
    if (
      typeof propOrValue !== 'object' ||
      propOrValue instanceof Date
    ) {
      resolve(propOrValue);
      return;
    }

    // Compose-mode Office.js wrapper.
    if (
      typeof propOrValue.getAsync === 'function'
    ) {
      propOrValue.getAsync(result => {
        resolve(
          result.status ===
            Office.AsyncResultStatus.Succeeded
            ? result.value
            : ''
        );
      });
    } else {
      resolve('');
    }
  });
}


/**
 * Get body as plain text.
 */
function getBodyText(item) {
  return new Promise(resolve => {
    if (!item.body) {
      resolve('');
      return;
    }

    item.body.getAsync(
      Office.CoercionType.Text,
      result => {
        resolve(
          result.status ===
            Office.AsyncResultStatus.Succeeded
            ? result.value || ''
            : ''
        );
      }
    );
  });
}


/**
 * Get the body as Markdown.
 *
 * Pulls HTML, cleans Outlook-specific markup and converts it using Turndown.
 * Falls back to plain text if Turndown is unavailable or conversion fails.
 */
function getBodyMarkdown(item) {
  return new Promise(resolve => {
    if (!item.body) {
      resolve('');
      return;
    }

    item.body.getAsync(
      Office.CoercionType.Html,
      result => {
        if (
          result.status !==
            Office.AsyncResultStatus.Succeeded ||
          !result.value
        ) {
          getBodyText(item)
            .then(text => resolve(removeTeamsInviteText(text)));
          return;
        }

        const html =
          flattenTableCellBlocks(
            promoteOutlookTableHeaders(
              removeTeamsInviteBlockHtml(
                cleanOutlookHtml(result.value)
              )
            )
          );

        if (
          typeof TurndownService === 'undefined'
        ) {
          const tmp =
            document.createElement('div');

          tmp.innerHTML = html;

          resolve(
            removeTeamsInviteText(tmp.textContent || '')
          );

          return;
        }

        try {
          const td =
            new TurndownService({
              headingStyle: 'atx',
              bulletListMarker: '-',
              codeBlockStyle: 'fenced',
              emDelimiter: '*',
            });

          // GFM plugin: tables, strikethrough, task lists.
          if (
            typeof turndownPluginGfm !==
            'undefined'
          ) {
            td.use(turndownPluginGfm.gfm);
          }

          // Inline Outlook cid: images cannot be resolved in Obsidian.
          td.addRule('strip-cid-images', {
            filter: node =>
              node.nodeName === 'IMG' &&
              (
                node.getAttribute('src') ||
                ''
              )
                .toLowerCase()
                .startsWith('cid:'),

            replacement: (_content, node) => {
              const alt =
                node.getAttribute('alt');

              return alt
                ? `*[${alt}]*`
                : '';
            },
          });

          const md =
            td.turndown(html)
              .replace(/\u00a0/g, ' ')
              .replace(/\n{3,}/g, '\n\n')
              .trim();

          resolve(removeTeamsInviteText(md));
        } catch (e) {
          console.warn(
            'Turndown conversion failed, using text fallback:',
            e.message
          );

          getBodyText(item)
            .then(text => resolve(removeTeamsInviteText(text)));
        }
      }
    );
  });
}


/**
 * Remove the complete Microsoft Teams invitation block from Outlook HTML.
 *
 * Primary strategy:
 *   1. Remove known Teams placeholder containers when Outlook provides them.
 *   2. Otherwise locate the Teams heading and remove the surrounding generated
 *      block up to the closing separator / organizer section.
 *   3. Finally remove any standalone Teams join/help links that remain.
 *
 * The original meeting description / agenda outside the Teams block is kept.
 */
function removeTeamsInviteBlockHtml(html) {
  const wrapper =
    document.createElement('div');

  wrapper.innerHTML = html || '';

  /*
   * Modern Outlook commonly wraps the generated Teams invitation in a
   * dedicated element. Remove those first when present.
   */
  const knownTeamsSelectors = [
    '[id*="MicrosoftTeamsMeetingPlaceholder"]',
    '[id*="TeamsMeetingPlaceholder"]',
    '[class*="MicrosoftTeamsMeetingPlaceholder"]',
    '[class*="TeamsMeetingPlaceholder"]',
  ];

  for (const selector of knownTeamsSelectors) {
    for (const node of wrapper.querySelectorAll(selector)) {
      node.remove();
    }
  }

  /*
   * Fallback for Outlook versions/locales where no useful Teams container ID
   * exists. Find a block whose text is the Teams heading and remove the
   * generated siblings that follow it.
   */
  const blockTags =
    new Set(['DIV', 'P', 'TABLE', 'TR', 'TD', 'LI']);

  const isTeamsHeading = value =>
    /^(?:Microsoft Teams(?:-möte| meeting))$/i
      .test(
        String(value || '')
          .replace(/\u00a0/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
      );

  const isSeparator = value =>
    /^[_\-—–]{20,}$/
      .test(
        String(value || '')
          .replace(/\u00a0/g, ' ')
          .replace(/\s+/g, '')
          .trim()
      );

  const isOrganizerLine = value =>
    /^(?:För organisatörer:|For organizers:)/i
      .test(
        String(value || '')
          .replace(/\u00a0/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
      );

  const allElements =
    Array.from(wrapper.querySelectorAll('*'));

  for (const element of allElements) {
    if (!element.isConnected) {
      continue;
    }

    const elementText =
      (element.textContent || '')
        .replace(/\u00a0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    if (!isTeamsHeading(elementText)) {
      continue;
    }

    /*
     * Walk upward to a sensible block-level node, but do not jump all the way
     * to the wrapper because that could remove the user's actual agenda.
     */
    let startNode = element;

    while (
      startNode.parentElement &&
      startNode.parentElement !== wrapper &&
      !blockTags.has(startNode.nodeName)
    ) {
      startNode = startNode.parentElement;
    }

    /*
     * Outlook usually places the generated Teams content in consecutive block
     * siblings. Remove from the Teams heading through the closing separator.
     *
     * If there is no separator, stop after the organizer line and a small
     * amount of generated trailing content.
     */
    let node = startNode;
    let organizerSeen = false;
    let removedCount = 0;

    while (node && node !== wrapper) {
      const next =
        node.nextElementSibling;

      const nodeText =
        (node.textContent || '')
          .replace(/\u00a0/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

      if (isOrganizerLine(nodeText)) {
        organizerSeen = true;
      }

      const separator =
        isSeparator(nodeText);

      node.remove();
      removedCount += 1;

      /*
       * The first separator may be the line immediately above the Teams
       * heading. Only treat a separator as the end once content has actually
       * been removed after the heading.
       */
      if (
        separator &&
        removedCount > 1
      ) {
        break;
      }

      /*
       * Some Outlook variants do not include the bottom separator as its own
       * element. Once the organizer line has been removed, stop if the next
       * sibling does not look like Teams boilerplate.
       */
      if (
        organizerSeen &&
        next
      ) {
        const nextText =
          (next.textContent || '')
            .replace(/\u00a0/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

        if (
          !isSeparator(nextText) &&
          !/^(?:Privacy and security|Sekretess och säkerhet)/i.test(nextText)
        ) {
          break;
        }
      }

      node = next;
    }
  }

  /*
   * Remove standalone Teams-related links that may remain after block removal.
   */
  for (const link of wrapper.querySelectorAll('a')) {
    const href =
      (link.getAttribute('href') || '').toLowerCase();

    const linkText =
      (link.textContent || '')
        .replace(/\u00a0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();

    if (
      href.includes('teams.microsoft.com') ||
      href.includes('teams.live.com') ||
      href.includes('aka.ms/jointeamsmeeting') ||
      linkText.includes('join microsoft teams') ||
      linkText.includes('join teams meeting') ||
      linkText === 'behöver du hjälp?' ||
      linkText === 'need help?'
    ) {
      const container =
        link.parentElement;

      const onlyContent =
        container &&
        container.textContent.trim() ===
          link.textContent.trim();

      if (
        onlyContent &&
        ['P', 'DIV', 'LI'].includes(container.nodeName)
      ) {
        container.remove();
      } else {
        link.remove();
      }
    }
  }

  return wrapper.innerHTML;
}


/**
 * Remove a Microsoft Teams invitation block from plain text / Markdown.
 *
 * This is intentionally kept as a second-stage fallback after Turndown.
 * HTML removal is the primary mechanism.
 */
function removeTeamsInviteText(text) {
  let result =
    String(text || '')
      .replace(/\r\n/g, '\n');

  /*
   * Main pattern: remove the Teams heading and everything up to the closing
   * separator line.
   */
  result = result.replace(
    /(?:^|\n)[ \t]*(?:[_\-—–]{10,})?[ \t]*\n?[ \t]*(?:\*\*)?Microsoft Teams(?:-möte| meeting)(?:\*\*)?[ \t]*\n[\s\S]*?(?=\n[ \t]*(?:[_\-—–]{10,})[ \t]*(?:\n|$)|$)/gi,
    '\n'
  );

  /*
   * Fallback when Turndown/Outlook has removed or transformed the separator.
   */
  result = result.replace(
    /(?:^|\n)[ \t]*(?:\*\*)?Microsoft Teams(?:-möte| meeting)(?:\*\*)?[ \t]*\n[\s\S]*?(?:För organisatörer:|For organizers:)[^\n]*(?:\n|$)/gi,
    '\n'
  );

  /*
   * Remove any remaining Teams URLs.
   */
  result = result.replace(
    /https?:\/\/(?:[\w-]+\.)?(?:teams\.microsoft\.com|teams\.live\.com)\/\S+/gi,
    ''
  );

  result = result.replace(
    /https?:\/\/aka\.ms\/JoinTeamsMeeting\S*/gi,
    ''
  );

  /*
   * Remove help-link text that may survive independently.
   */
  result = result.replace(
    /(?:^|\n)[ \t]*(?:\[[^\]]*\]\([^)]+\)|(?:Behöver du hjälp\?|Need help\?))[ \t]*(?:\|)?[ \t]*(?=\n|$)/gi,
    '\n'
  );

  /*
   * Clean up separator-only lines left by Outlook.
   */
  result = result.replace(
    /(?:^|\n)[ \t]*[_\-—–]{20,}[ \t]*(?=\n|$)/g,
    ''
  );

  return result
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}


/**
 * Outlook often represents a table header as bold <td> cells rather than
 * actual <th> elements. Promote the first row to <th> when no header exists
 * so Turndown's GFM table converter can recognize the table.
 */
function promoteOutlookTableHeaders(html) {
  const wrapper =
    document.createElement('div');

  wrapper.innerHTML = html;

  for (
    const table
    of wrapper.querySelectorAll('table')
  ) {
    if (
      table.querySelector('th') ||
      table.querySelector('thead')
    ) {
      continue;
    }

    const firstRow =
      table.querySelector('tr');

    if (!firstRow) {
      continue;
    }

    const tdCells =
      Array.from(firstRow.children)
        .filter(
          cell => cell.nodeName === 'TD'
        );

    if (tdCells.length === 0) {
      continue;
    }

    for (const td of tdCells) {
      const th =
        document.createElement('th');

      for (
        const attr
        of Array.from(td.attributes)
      ) {
        th.setAttribute(
          attr.name,
          attr.value
        );
      }

      while (td.firstChild) {
        th.appendChild(td.firstChild);
      }

      td.parentNode.replaceChild(
        th,
        td
      );
    }
  }

  return wrapper.innerHTML;
}


/**
 * Markdown tables require cell content to remain on one line. Outlook wraps
 * content in <p>/<div>, so unwrap block elements inside table cells.
 */
function flattenTableCellBlocks(html) {
  const wrapper =
    document.createElement('div');

  wrapper.innerHTML = html;

  for (
    const cell
    of wrapper.querySelectorAll('td, th')
  ) {
    let block;

    while (
      (
        block =
          cell.querySelector('p, div')
      )
    ) {
      const parent =
        block.parentNode;

      while (block.firstChild) {
        parent.insertBefore(
          block.firstChild,
          block
        );
      }

      parent.removeChild(block);
    }
  }

  return wrapper.innerHTML;
}


/**
 * Strip Office/Outlook HTML cruft that causes noisy Markdown conversion.
 */
function cleanOutlookHtml(html) {
  return html
    // MSO conditional comments.
    .replace(
      /<!--\[if[^>]*?\]>[\s\S]*?<!\[endif\]-->/gi,
      ''
    )

    // Outlook <o:p> elements.
    .replace(
      /<o:p[^>]*>[\s\S]*?<\/o:p>/gi,
      ''
    )
    .replace(
      /<o:p[^>]*\/>/gi,
      ''
    )

    // Remove complete style blocks before stripping remaining head tags.
    .replace(
      /<style[\s\S]*?<\/style>/gi,
      ''
    )

    // Remaining metadata/head leftovers.
    .replace(
      /<\/?(?:meta|link|style)\b[^>]*>/gi,
      ''
    );
}


/* ═══════════════════════════════════════════════════════════════════════════
   ATTENDEE EXTRACTION — OFFICE.JS
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Extract attendees using Office.js.
 *
 * When appointmentResponse is populated, attendees are grouped by RSVP.
 * Otherwise they appear under "No Response".
 */
function getAttendeesViaOfficeJs(item) {
  return new Promise(async resolve => {
    const result = {
      'Accepted':    new Map(),
      'Tentative':   new Map(),
      'Declined':    new Map(),
      'No Response': new Map(),
    };

    const RESPONSE_MAP = {
      accepted:  'Accepted',
      tentative: 'Tentative',
      declined:  'Declined',
      none:      'No Response',
      organizer: null,
    };

    async function fetchList(prop) {
      const value =
        item[prop];

      if (!value) {
        return [];
      }

      // Read mode.
      if (Array.isArray(value)) {
        return value;
      }

      // Compose mode.
      if (
        typeof value.getAsync === 'function'
      ) {
        return await new Promise(res => {
          value.getAsync(r => {
            res(
              r.status ===
                Office.AsyncResultStatus.Succeeded
                ? r.value || []
                : []
            );
          });
        });
      }

      return [];
    }

    const required =
      await fetchList('requiredAttendees');

    const optional =
      await fetchList('optionalAttendees');

    for (
      const attendee
      of [...required, ...optional]
    ) {
      const name =
        attendee.displayName ||
        attendee.emailAddress ||
        '';

      const email =
        (attendee.emailAddress || '')
          .toLowerCase();

      if (!name) {
        continue;
      }

      const responseType =
        (
          attendee.appointmentResponse ||
          'none'
        )
          .toString()
          .toLowerCase();

      const bucket =
        Object.prototype.hasOwnProperty.call(
          RESPONSE_MAP,
          responseType
        )
          ? RESPONSE_MAP[responseType]
          : 'No Response';

      // Organizer is captured separately from item.organizer.
      if (bucket === null) {
        continue;
      }

      result[bucket].set(
        name,
        email
      );
    }

    resolve(result);
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
   EWS — ATTENDEES WITH RSVP STATUS
   ═══════════════════════════════════════════════════════════════════════════
   Preserved from the upstream implementation.

   Requires ReadWriteMailbox permission in the existing manifest.
   If EWS fails, extractMeetingDetails() falls back to Office.js.
   ═══════════════════════════════════════════════════════════════════════════ */

function getAttendeesViaEws(itemId) {
  return new Promise((resolve, reject) => {
    // Some Outlook clients expose a REST-formatted item ID.
    // EWS needs its own ID format.
    let ewsId = itemId;

    try {
      const converted =
        Office.context.mailbox.convertToEwsId(
          itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );

      if (converted) {
        ewsId = converted;
      }
    } catch (_) {
      // Already EWS format or conversion unavailable.
    }

    const soap =
`<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:RequiredAttendees"/>
          <t:FieldURI FieldURI="calendar:OptionalAttendees"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${ewsId}"/>
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(
      soap,
      result => {
        if (
          result.status !==
          Office.AsyncResultStatus.Succeeded
        ) {
          const err =
            result.error || {};

          const code =
            err.code != null
              ? ` code=${err.code}`
              : '';

          const name =
            err.name
              ? ` name=${err.name}`
              : '';

          const body =
            typeof result.value === 'string' &&
            result.value.length > 0
              ? (
                  ' body=' +
                  result.value
                    .replace(/\s+/g, ' ')
                    .slice(0, 400)
                )
              : '';

          reject(
            new Error(
              (
                err.message ||
                'EWS request failed'
              ) +
              code +
              name +
              body
            )
          );

          return;
        }

        try {
          resolve(
            parseEwsAttendees(
              result.value
            )
          );
        } catch (e) {
          reject(e);
        }
      }
    );
  });
}


/**
 * Parse EWS GetItem XML and return attendees grouped by RSVP status.
 *
 * EWS ResponseType values include:
 *   None
 *   Organizer
 *   Tentative
 *   Accept
 *   Decline
 *   NoResponseReceived
 */
function parseEwsAttendees(xmlString) {
  const parser =
    new DOMParser();

  const doc =
    parser.parseFromString(
      xmlString,
      'text/xml'
    );

  const T =
    'http://schemas.microsoft.com/exchange/services/2006/types';

  const M =
    'http://schemas.microsoft.com/exchange/services/2006/messages';


  /* ── Surface EWS errors ───────────────────────────────────────────────── */

  const responseMsgs =
    doc.getElementsByTagNameNS(
      M,
      'GetItemResponseMessage'
    );

  if (responseMsgs.length > 0) {
    const responseClass =
      responseMsgs[0]
        .getAttribute('ResponseClass');

    if (responseClass === 'Error') {
      const code =
        responseMsgs[0]
          .getElementsByTagNameNS(
            M,
            'ResponseCode'
          )[0]
          ?.textContent ||
        '';

      throw new Error(
        `EWS error: ${code}`
      );
    }
  }


  /* ── RSVP buckets ─────────────────────────────────────────────────────── */

  const result = {
    'Accepted':    new Map(),
    'Tentative':   new Map(),
    'Declined':    new Map(),
    'No Response': new Map(),
  };

  const RSVP_MAP = {
    accept:             'Accepted',
    tentative:          'Tentative',
    decline:            'Declined',
    none:               'No Response',
    noresponsereceived: 'No Response',
    organizer:          null,
  };


  function processAttendeeGroup(tagName) {
    for (
      const group
      of doc.getElementsByTagNameNS(
        T,
        tagName
      )
    ) {
      for (
        const attendeeEl
        of group.getElementsByTagNameNS(
          T,
          'Attendee'
        )
      ) {
        const name =
          attendeeEl
            .getElementsByTagNameNS(
              T,
              'Name'
            )[0]
            ?.textContent
            ?.trim() ||
          '';

        const email =
          (
            attendeeEl
              .getElementsByTagNameNS(
                T,
                'EmailAddress'
              )[0]
              ?.textContent
              ?.trim() ||
            ''
          )
            .toLowerCase();

        const responseType =
          (
            attendeeEl
              .getElementsByTagNameNS(
                T,
                'ResponseType'
              )[0]
              ?.textContent
              ?.trim() ||
            'none'
          )
            .toLowerCase();

        if (!name) {
          continue;
        }

        const bucket =
          Object.prototype.hasOwnProperty.call(
            RSVP_MAP,
            responseType
          )
            ? RSVP_MAP[responseType]
            : 'No Response';

        // Organizer is captured separately.
        if (bucket === null) {
          continue;
        }

        result[bucket].set(
          name,
          email
        );
      }
    }
  }

  processAttendeeGroup(
    'RequiredAttendees'
  );

  processAttendeeGroup(
    'OptionalAttendees'
  );

  return result;
}