'use strict';

/* ═══════════════════════════════════════════════════════════════════════════
   Outlook to Obsidian — Office Add-in task pane
   ═══════════════════════════════════════════════════════════════════════════
   Uses Office.js (Office.context.mailbox.item) to read meeting/email data.

   This version separates Outlook data extraction from note presentation:
     templates/event-template.md
     templates/mail-template.md

   Template placeholders use the form {{variable_name}}.

   RSVP status is fetched via EWS (makeEwsRequestAsync) when possible.
   If EWS is unavailable, the add-in falls back to Office.js attendee data.

   Meeting body processing pipeline:

     Outlook HTML
       ↓
     cleanOutlookHtml()
       ↓
     removeTeamsInviteBlockHtml()
       ↓
     normalizeOutlookListNesting()
       ↓
     Turndown
       ↓
     removeTeamsInviteText()
       ↓
     compactMarkdown()

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


/**
 * Returns 'message' or 'meeting' based on the current Outlook item.
 */
function getCurrentItemType() {
  try {
    const t =
      Office.context.mailbox.item?.itemType;

    return t === Office.MailboxEnums.ItemType.Message
      ? 'message'
      : 'meeting';
  } catch (_) {
    return 'meeting';
  }
}


/**
 * Update button label and header to match the current context.
 */
function applyContextLabels() {
  const isMessage =
    getCurrentItemType() === 'message';

  const btn =
    document.getElementById('extractBtn');

  if (btn) {
    btn.innerHTML =
      isMessage
        ? '<span>📧</span><span>Extract Email Details</span>'
        : '<span>📅</span><span>Extract Meeting Details</span>';
  }

  const headerIcon =
    document.getElementById('headerIcon');

  if (headerIcon) {
    headerIcon.textContent =
      isMessage ? '📧' : '📅';
  }

  const headerTitle =
    document.getElementById('headerTitle');

  if (headerTitle) {
    headerTitle.textContent =
      isMessage
        ? 'Send Email to Obsidian'
        : 'Send Meeting to Obsidian';
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   UI BINDING
   ═══════════════════════════════════════════════════════════════════════════ */

function bindUI() {
  document
    .getElementById('extractBtn')
    .addEventListener(
      'click',
      onExtract
    );

  document
    .getElementById('openObsidianBtn')
    .addEventListener(
      'click',
      onOpenObsidian
    );

  document
    .getElementById('copyBtn')
    .addEventListener(
      'click',
      onCopy
    );

  document
    .getElementById('settingsBtn')
    .addEventListener(
      'click',
      () => showPanel('settings')
    );

  document
    .getElementById('backBtn')
    .addEventListener(
      'click',
      () => showPanel('main')
    );

  document
    .getElementById('saveSettingsBtn')
    .addEventListener(
      'click',
      onSaveSettings
    );

  document
    .querySelectorAll('.preset-btn')
    .forEach(btn => {
      btn.addEventListener(
        'click',
        () => {
          const target =
            btn.dataset.target;

          const value =
            btn.dataset.value;

          if (
            target &&
            value
          ) {
            document
              .getElementById(target)
              .value = value;

            if (
              target === 'vaultName' ||
              target === 'folderPath'
            ) {
              updateFolderPreview();
            }
          }
        }
      );
    });

  document
    .getElementById('vaultName')
    .addEventListener(
      'input',
      updateFolderPreview
    );

  document
    .getElementById('folderPath')
    .addEventListener(
      'input',
      updateFolderPreview
    );
}


/* ═══════════════════════════════════════════════════════════════════════════
   PANEL SWITCHING
   ═══════════════════════════════════════════════════════════════════════════ */

function showPanel(name) {
  const isMain =
    name === 'main';

  document
    .getElementById('mainPanel')
    .style.display =
      isMain ? 'flex' : 'none';

  document
    .getElementById('settingsPanel')
    .style.display =
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
    const rs =
      Office.context.roamingSettings;

    for (
      const key
      of Object.keys(DEFAULT_SETTINGS)
    ) {
      const value =
        rs.get(key);

      if (
        value !== undefined &&
        value !== null
      ) {
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
    document
      .getElementById('vaultName')
      .value
      .trim();

  const folderPath =
    document
      .getElementById('folderPath')
      .value
      .trim();

  const emailFolderPath =
    document
      .getElementById('emailFolderPath')
      .value
      .trim();

  const internalDomain =
    document
      .getElementById('internalDomain')
      .value
      .trim();

  if (!vaultName) {
    showStatus(
      'Please enter a vault name.',
      'error'
    );
    return;
  }

  if (!folderPath) {
    showStatus(
      'Please enter a meeting notes folder.',
      'error'
    );
    return;
  }

  if (!emailFolderPath) {
    showStatus(
      'Please enter an email notes folder.',
      'error'
    );
    return;
  }

  settings.vaultName =
    vaultName;

  settings.folderPath =
    folderPath;

  settings.emailFolderPath =
    emailFolderPath;

  settings.internalDomain =
    internalDomain ||
    DEFAULT_SETTINGS.internalDomain;

  settings.copyToClipboard =
    document
      .getElementById('copyToClipboard')
      .checked;

  settings.includeExternal =
    document
      .getElementById('includeExternal')
      .checked;

  try {
    const rs =
      Office.context.roamingSettings;

    for (
      const [key, value]
      of Object.entries(settings)
    ) {
      rs.set(
        key,
        value
      );
    }

    rs.saveAsync(result => {
      if (
        result.status ===
        Office.AsyncResultStatus.Succeeded
      ) {
        updateConfigDisplay();
        showPanel('main');

        showStatus(
          'Settings saved.',
          'success'
        );
      } else {
        showStatus(
          'Could not save settings: ' +
            (
              result.error?.message ||
              'unknown error'
            ),
          'error'
        );
      }
    });
  } catch (e) {
    showStatus(
      'Error saving settings: ' +
        e.message,
      'error'
    );
  }
}


function populateSettingsForm() {
  document
    .getElementById('vaultName')
    .value =
      settings.vaultName;

  document
    .getElementById('folderPath')
    .value =
      settings.folderPath;

  document
    .getElementById('emailFolderPath')
    .value =
      settings.emailFolderPath;

  document
    .getElementById('internalDomain')
    .value =
      settings.internalDomain;

  document
    .getElementById('copyToClipboard')
    .checked =
      settings.copyToClipboard;

  document
    .getElementById('includeExternal')
    .checked =
      settings.includeExternal;
}


function updateConfigDisplay() {
  const isMessage =
    getCurrentItemType() === 'message';

  const activePath =
    isMessage
      ? settings.emailFolderPath
      : settings.folderPath;

  const ok =
    !!(
      settings.vaultName &&
      activePath
    );

  document
    .getElementById('configWarning')
    .style.display =
      ok ? 'none' : 'block';

  document
    .getElementById('vaultDisplay')
    .textContent =
      settings.vaultName ||
      'Not set';

  document
    .getElementById('pathDisplay')
    .textContent =
      activePath ||
      'Not set';

  document
    .getElementById('extractBtn')
    .disabled =
      !ok;
}


function updateFolderPreview() {
  const vault =
    document
      .getElementById('vaultName')
      .value ||
    'YourVault';

  const folder =
    document
      .getElementById('folderPath')
      .value ||
    'Meeting Notes';

  const now =
    new Date();

  const y =
    now.getFullYear();

  const m =
    String(
      now.getMonth() + 1
    ).padStart(
      2,
      '0'
    );

  document
    .getElementById('folderPreview')
    .textContent =
      `📁 ${vault}/\n` +
      `  📁 ${folder}/\n` +
      `    📁 ${y}/\n` +
      `      📁 ${m}/\n` +
      `        📄 ${y}-${m}-15 Example Meeting.md`;
}


/* ═══════════════════════════════════════════════════════════════════════════
   TEMPLATE SYSTEM
   ═══════════════════════════════════════════════════════════════════════════ */

async function loadNoteTemplate(filename) {
  const response =
    await fetch(
      `./templates/${filename}`,
      {
        cache: 'no-store',
      }
    );

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
 */
function renderNoteTemplate(
  template,
  values
) {
  return template.replace(
    /\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g,
    (
      _match,
      key
    ) => {
      if (
        !Object.prototype
          .hasOwnProperty
          .call(
            values,
            key
          )
      ) {
        console.warn(
          `Unknown template variable: ${key}`
        );

        return '';
      }

      return values[key] ?? '';
    }
  );
}


/**
 * Quote a scalar safely for YAML frontmatter.
 */
function yamlQuote(value) {
  return `"${String(value ?? '')
    .replace(
      /\\/g,
      '\\\\'
    )
    .replace(
      /"/g,
      '\\"'
    )
    .replace(
      /\r?\n/g,
      '\\n'
    )}"`;
}


/**
 * Format an array as a YAML list property.
 */
function yamlListProperty(
  propertyName,
  values
) {
  const unique =
    Array.from(
      new Set(
        (values || [])
          .filter(Boolean)
      )
    );

  if (
    unique.length === 0
  ) {
    return `${propertyName}: []`;
  }

  return [
    `${propertyName}:`,
    ...unique.map(
      value =>
        `  - ${yamlQuote(value)}`
    ),
  ].join('\n');
}


/* ═══════════════════════════════════════════════════════════════════════════
   EXTRACT BUTTON HANDLER
   ═══════════════════════════════════════════════════════════════════════════ */

async function onExtract() {
  const btn =
    document
      .getElementById('extractBtn');

  btn.disabled = true;

  document
    .getElementById('details')
    .style.display =
      'none';

  document
    .getElementById('actions')
    .style.display =
      'none';

  const isMessage =
    getCurrentItemType() ===
    'message';

  showStatus(
    isMessage
      ? 'Extracting email…'
      : 'Extracting meeting details…',
    'info'
  );

  try {
    extractedData =
      isMessage
        ? await extractEmailDetails()
        : await extractMeetingDetails();

    if (
      settings.includeExternal &&
      extractedData.hasExternalAttendees
    ) {
      extractedData.title =
        '[EXTERNAL] ' +
        extractedData.title;

      extractedData.note =
        extractedData.note.replace(
          /^(# ).+$/m,
          `$1${extractedData.title}`
        );
    }

    renderDetails(
      extractedData
    );

    showStatus(
      isMessage
        ? 'Email extracted successfully.'
        : 'Meeting details extracted successfully.',
      'success'
    );

    if (
      settings.copyToClipboard
    ) {
      await copyText(
        extractedData.note
      );
    }

    document
      .getElementById('details')
      .style.display =
        'block';

    document
      .getElementById('actions')
      .style.display =
        'flex';

  } catch (err) {
    showStatus(
      'Error: ' +
        err.message,
      'error'
    );

  } finally {
    const activePath =
      isMessage
        ? settings.emailFolderPath
        : settings.folderPath;

    btn.disabled =
      !(
        settings.vaultName &&
        activePath
      );
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   DETAILS PREVIEW
   ═══════════════════════════════════════════════════════════════════════════ */

function renderDetails(data) {
  const truncate =
    (
      value,
      maxLength
    ) => {
      const s =
        String(
          value || ''
        );

      return s.length >
        maxLength
        ? s.substring(
            0,
            maxLength
          ) + '…'
        : s;
    };

  const isEmail =
    data.kind === 'email';

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

  if (
    !isEmail &&
    data.location
  ) {
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

  if (
    data.hasExternalAttendees
  ) {
    rows.push(`
      <div class="detail-row">
        <span class="detail-label">External</span>
        <span class="detail-value external">Yes</span>
      </div>
    `);
  }

  document
    .getElementById('details')
    .innerHTML =
      rows.join('');
}


/* ═══════════════════════════════════════════════════════════════════════════
   OPEN IN OBSIDIAN
   ═══════════════════════════════════════════════════════════════════════════ */

function onOpenObsidian() {
  if (!extractedData) {
    return;
  }

  const uri =
    buildObsidianUri(
      extractedData
    );

  const link =
    document
      .getElementById('obsidianLink');

  link.href = uri;
  link.click();
}


function buildObsidianUri(data) {
  const basePath =
    data.kind === 'email'
      ? settings.emailFolderPath
      : settings.folderPath;

  const folder =
    `${basePath}/` +
    `${data.meetingDate.year}/` +
    `${data.meetingDate.month}`;

  const safeTitle =
    data.title.replace(
      /[<>:"/\\|?*]/g,
      '-'
    );

  const file =
    `${folder}/` +
    `${data.meetingDate.full} ` +
    `${safeTitle}`;

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

  await copyText(
    extractedData.note
  );

  showStatus(
    'Note copied to clipboard.',
    'success'
  );
}


async function copyText(text) {
  try {
    await navigator.clipboard
      .writeText(text);

  } catch (_) {
    const ta =
      document
        .createElement('textarea');

    ta.value =
      text;

    ta.style.position =
      'fixed';

    ta.style.opacity =
      '0';

    document.body
      .appendChild(ta);

    ta.focus();
    ta.select();

    document.execCommand(
      'copy'
    );

    document.body
      .removeChild(ta);
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
   UI HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */

function showStatus(
  message,
  type
) {
  const el =
    document
      .getElementById('status');

  el.textContent =
    message;

  el.className =
    `status ${type}`;

  el.style.display =
    'block';
}


function esc(value) {
  return String(
    value || ''
  )
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    );
}


/* ═══════════════════════════════════════════════════════════════════════════
   CORE EXTRACTION — MEETING
   ═══════════════════════════════════════════════════════════════════════════ */

async function extractMeetingDetails() {
  const item =
    Office.context.mailbox.item;

  const isReadMode =
    !!item.itemId;


  /* ── Subject ──────────────────────────────────────────────────────────── */

  const title =
    (
      await getItemProperty(
        item.subject
      )
    ) ||
    'Untitled Meeting';


  /* ── Start / end ──────────────────────────────────────────────────────── */

  const startDate =
    await getItemProperty(
      item.start
    );

  const endDate =
    await getItemProperty(
      item.end
    );

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
      String(
        startDate.getMonth() + 1
      ).padStart(
        2,
        '0'
      );

    const d =
      String(
        startDate.getDate()
      ).padStart(
        2,
        '0'
      );

    meetingDate = {
      year:
        String(y),

      month:
        mo,

      day:
        d,

      full:
        `${y}-${mo}-${d}`,
    };

    const timeFormat = {
      hour:
        '2-digit',

      minute:
        '2-digit',
    };

    startTime =
      startDate
        .toLocaleTimeString(
          [],
          timeFormat
        );

    endTime =
      endDate instanceof Date &&
      !isNaN(endDate)
        ? endDate
            .toLocaleTimeString(
              [],
              timeFormat
            )
        : '';

    time =
      `${startDate.toLocaleDateString()} ${startTime}` +
      (
        endTime
          ? ` – ${endTime}`
          : ''
      );
  }

  if (!meetingDate) {
    const today =
      new Date();

    const y =
      today.getFullYear();

    const mo =
      String(
        today.getMonth() + 1
      ).padStart(
        2,
        '0'
      );

    const d =
      String(
        today.getDate()
      ).padStart(
        2,
        '0'
      );

    meetingDate = {
      year:
        String(y),

      month:
        mo,

      day:
        d,

      full:
        `${y}-${mo}-${d}`,
    };
  }


  /* ── Location ─────────────────────────────────────────────────────────── */

  const location =
    (
      await getItemProperty(
        item.location
      )
    ) ||
    '';


  /* ── Current Outlook user ─────────────────────────────────────────────── */

  const currentUserName =
    (
      Office.context
        .mailbox
        .userProfile
        ?.displayName ||
      ''
    )
      .trim();

  const currentUserEmail =
    (
      Office.context
        .mailbox
        .userProfile
        ?.emailAddress ||
      ''
    )
      .trim()
      .toLowerCase();


  /* ── Organizer ────────────────────────────────────────────────────────── */

  let organizer = '';
  let organizerEmail = '';
  let organizerIsSelf = false;

  if (
    isReadMode &&
    item.organizer
  ) {
    organizer =
      (
        item.organizer
          .displayName ||
        ''
      )
        .trim();

    organizerEmail =
      (
        item.organizer
          .emailAddress ||
        ''
      )
        .trim()
        .toLowerCase();

    organizerIsSelf =
      !!(
        currentUserEmail &&
        organizerEmail &&
        currentUserEmail ===
          organizerEmail
      );
  }


  /*
   * Compose mode normally does not expose item.organizer.
   * The mailbox user creating the appointment is therefore used
   * as the organizer fallback.
   */
  if (
    !isReadMode &&
    currentUserEmail
  ) {
    organizer =
      currentUserName ||
      currentUserEmail;

    organizerEmail =
      currentUserEmail;

    organizerIsSelf =
      true;
  }


  /*
   * If Outlook identifies the email as self but for some reason
   * does not return a display name, use the mailbox profile name.
   */
  if (
    organizerIsSelf &&
    !organizer
  ) {
    organizer =
      currentUserName ||
      currentUserEmail;
  }


  /*
   * Presentation value for the note.
   *
   * Example:
   *   Jonatan Hellborg (self)
   */
  const organizerDisplay =
    organizerIsSelf
      ? (
          `${organizer || currentUserName || currentUserEmail} (self)`
        )
      : organizer;


  /* ── Body / description ───────────────────────────────────────────────── */

  const body =
    await getBodyMarkdown(
      item
    );


  /* ── Attendees with RSVP ──────────────────────────────────────────────── */

  let attendeesByStatus = {
    'Accepted':
      new Map(),

    'Tentative':
      new Map(),

    'Declined':
      new Map(),

    'No Response':
      new Map(),
  };

  let ewsErrorMessage =
    '';

  if (
    isReadMode &&
    item.itemId
  ) {
    try {
      attendeesByStatus =
        await getAttendeesViaEws(
          item.itemId
        );

    } catch (ewsErr) {
      console.warn(
        'EWS attendee fetch failed, using Office.js fallback:',
        ewsErr.message
      );

      ewsErrorMessage =
        ewsErr.message ||
        'unknown';

      attendeesByStatus =
        await getAttendeesViaOfficeJs(
          item
        );
    }

  } else {
    attendeesByStatus =
      await getAttendeesViaOfficeJs(
        item
      );
  }


  /* ── External attendee detection ──────────────────────────────────────── */

  const internalSuffix =
    '@' +
    (
      settings.internalDomain ||
      DEFAULT_SETTINGS.internalDomain
    )
      .toLowerCase()
      .replace(
        /^@/,
        ''
      );

  let hasExternalAttendees =
    organizerEmail
      ? !organizerEmail
          .endsWith(
            internalSuffix
          )
      : false;

  if (
    !hasExternalAttendees
  ) {
    for (
      const attendees
      of Object.values(
        attendeesByStatus
      )
    ) {
      for (
        const email
        of attendees.values()
      ) {
        if (
          email &&
          !email.endsWith(
            internalSuffix
          )
        ) {
          hasExternalAttendees =
            true;

          break;
        }
      }

      if (
        hasExternalAttendees
      ) {
        break;
      }
    }
  }


  /* ── Format attendee information ──────────────────────────────────────── */

  const attendeesFormatted =
    [];

  const allAttendeeNames =
    [];

  let totalAttendees =
    organizer
      ? 1
      : 0;

  if (
    organizerDisplay
  ) {
    attendeesFormatted.push(
      `**Organizer:** ${organizerDisplay}`
    );

    /*
     * Keep the actual person's name in the YAML attendee list,
     * without "(self)" as that is presentation metadata rather
     * than part of the person's name.
     */
    allAttendeeNames.push(
      organizer ||
      currentUserName ||
      currentUserEmail
    );
  }

  for (
    const status
    of [
      'Accepted',
      'Tentative',
      'No Response',
      'Declined',
    ]
  ) {
    const list =
      Array.from(
        attendeesByStatus[
          status
        ].keys()
      )
        .sort();

    if (
      list.length > 0
    ) {
      totalAttendees +=
        list.length;

      attendeesFormatted.push(
        `**${status} (${list.length}):** ` +
        list.join(', ')
      );

      allAttendeeNames.push(
        ...list
      );
    }
  }

  const attendeesStr =
    attendeesFormatted
      .join('\n\n') ||
    'No attendees found';

  const attendeesFrontmatter =
    yamlListProperty(
      'attendees',
      allAttendeeNames
    );


  /* ── Render event-template.md ─────────────────────────────────────────── */

  const template =
    await loadNoteTemplate(
      'event-template.md'
    );

  const note =
    renderNoteTemplate(
      template,
      {
        // General
        title,

        // Date
        date:
          meetingDate.full,

        year:
          meetingDate.year,

        month:
          meetingDate.month,

        day:
          meetingDate.day,

        // Time
        time:
          time ||
          'Not specified',

        start_time:
          startTime,

        end_time:
          endTime,

        start_time_yaml:
          yamlQuote(
            startTime
          ),

        end_time_yaml:
          yamlQuote(
            endTime
          ),

        // Location
        location,

        location_yaml:
          yamlQuote(
            location
          ),

        // Organizer
        organizer:
          organizerDisplay,

        organizer_name:
          organizer,

        organizer_email:
          organizerEmail,

        organizer_is_self:
          String(
            organizerIsSelf
          ),

        organizer_yaml:
          yamlQuote(
            organizerDisplay
          ),

        organizer_name_yaml:
          yamlQuote(
            organizer
          ),

        organizer_email_yaml:
          yamlQuote(
            organizerEmail
          ),

        // Current user
        current_user_name:
          currentUserName,

        current_user_email:
          currentUserEmail,

        current_user_name_yaml:
          yamlQuote(
            currentUserName
          ),

        current_user_email_yaml:
          yamlQuote(
            currentUserEmail
          ),

        // Attendees
        attendees:
          attendeesStr,

        attendees_frontmatter:
          attendeesFrontmatter,

        attendee_count:
          String(
            totalAttendees
          ),

        // RSVP lists
        accepted_attendees:
          Array.from(
            attendeesByStatus[
              'Accepted'
            ].keys()
          )
            .sort()
            .join(', '),

        tentative_attendees:
          Array.from(
            attendeesByStatus[
              'Tentative'
            ].keys()
          )
            .sort()
            .join(', '),

        no_response_attendees:
          Array.from(
            attendeesByStatus[
              'No Response'
            ].keys()
          )
            .sort()
            .join(', '),

        declined_attendees:
          Array.from(
            attendeesByStatus[
              'Declined'
            ].keys()
          )
            .sort()
            .join(', '),

        // Meeting classification
        external:
          String(
            hasExternalAttendees
          ),

        external_line:
          hasExternalAttendees
            ? '**External Meeting:** Yes'
            : '',

        // Outlook meeting description / agenda
        body:
          body.trim() ||
          '',
      }
    );


  /* ── Return extracted meeting data ────────────────────────────────────── */

  const safeTitle =
    title.replace(
      /[<>:"/\\|?*]/g,
      '-'
    );

  return {
    kind:
      'meeting',

    title,

    time,

    startTime,

    endTime,

    location,

    organizer:
      organizerDisplay,

    organizerName:
      organizer,

    organizerEmail,

    organizerIsSelf,

    currentUserName,

    currentUserEmail,

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
    (
      await getItemProperty(
        item.subject
      )
    ) ||
    'No Subject';


  /* ── Sent date ────────────────────────────────────────────────────────── */

  const sentDate =
    item.dateTimeCreated
      instanceof Date &&
    !isNaN(
      item.dateTimeCreated
    )
      ? item.dateTimeCreated
      : (
          item.dateTimeModified
            instanceof Date &&
          !isNaN(
            item.dateTimeModified
          )
            ? item.dateTimeModified
            : new Date()
        );

  const y =
    sentDate.getFullYear();

  const mo =
    String(
      sentDate.getMonth() + 1
    ).padStart(
      2,
      '0'
    );

  const d =
    String(
      sentDate.getDate()
    ).padStart(
      2,
      '0'
    );

  const meetingDate = {
    year:
      String(y),

    month:
      mo,

    day:
      d,

    full:
      `${y}-${mo}-${d}`,
  };

  const sentTime =
    sentDate
      .toLocaleTimeString(
        [],
        {
          hour:
            '2-digit',

          minute:
            '2-digit',
        }
      );

  const time =
    `${sentDate.toLocaleDateString()} ${sentTime}`;


  /* ── From / sender ────────────────────────────────────────────────────── */

  const fromObj =
    item.from ||
    item.sender ||
    {};

  const fromName =
    fromObj.displayName ||
    fromObj.emailAddress ||
    '';

  const fromEmail =
    (
      fromObj.emailAddress ||
      ''
    )
      .toLowerCase();


  /* ── To / Cc ──────────────────────────────────────────────────────────── */

  const toList =
    Array.isArray(
      item.to
    )
      ? item.to
      : [];

  const ccList =
    Array.isArray(
      item.cc
    )
      ? item.cc
      : [];


  function formatPerson(
    person
  ) {
    const name =
      person.displayName ||
      person.emailAddress ||
      '';

    const email =
      person.emailAddress ||
      '';

    return (
      email &&
      name !== email
        ? `${name} <${email}>`
        : name
    );
  }


  const toFormatted =
    toList
      .map(
        formatPerson
      )
      .filter(
        Boolean
      );

  const ccFormatted =
    ccList
      .map(
        formatPerson
      )
      .filter(
        Boolean
      );


  /* ── Body ─────────────────────────────────────────────────────────────── */

  const body =
    await getBodyMarkdown(
      item
    );


  /* ── External detection ───────────────────────────────────────────────── */

  const internalSuffix =
    '@' +
    (
      settings.internalDomain ||
      DEFAULT_SETTINGS.internalDomain
    )
      .toLowerCase()
      .replace(
        /^@/,
        ''
      );

  let hasExternalAttendees =
    !!(
      fromEmail &&
      !fromEmail.endsWith(
        internalSuffix
      )
    );

  if (
    !hasExternalAttendees
  ) {
    for (
      const person
      of [
        ...toList,
        ...ccList,
      ]
    ) {
      const email =
        (
          person.emailAddress ||
          ''
        )
          .toLowerCase();

      if (
        email &&
        !email.endsWith(
          internalSuffix
        )
      ) {
        hasExternalAttendees =
          true;

        break;
      }
    }
  }

  const totalAttendees =
    toList.length +
    ccList.length;


  /* ── Sender display value ─────────────────────────────────────────────── */

  const fromLine =
    fromName
      ? (
          fromEmail &&
          fromName !==
            fromEmail
            ? `${fromName} <${fromEmail}>`
            : fromName
        )
      : 'Unknown sender';


  /* ── Render mail-template.md ──────────────────────────────────────────── */

  const template =
    await loadNoteTemplate(
      'mail-template.md'
    );

  const note =
    renderNoteTemplate(
      template,
      {
        title,

        date:
          meetingDate.full,

        year:
          meetingDate.year,

        month:
          meetingDate.month,

        day:
          meetingDate.day,

        time,

        sent_time:
          sentTime,

        from:
          fromLine,

        from_name:
          fromName,

        from_email:
          fromEmail,

        from_yaml:
          yamlQuote(
            fromLine
          ),

        from_name_yaml:
          yamlQuote(
            fromName
          ),

        from_email_yaml:
          yamlQuote(
            fromEmail
          ),

        to:
          toFormatted
            .join(', '),

        cc:
          ccFormatted
            .join(', '),

        to_line:
          toFormatted.length
            ? (
                `**To (${toFormatted.length}):** ` +
                toFormatted.join(', ')
              )
            : '',

        cc_line:
          ccFormatted.length
            ? (
                `**Cc (${ccFormatted.length}):** ` +
                ccFormatted.join(', ')
              )
            : '',

        to_frontmatter:
          yamlListProperty(
            'to',
            toFormatted
          ),

        cc_frontmatter:
          yamlListProperty(
            'cc',
            ccFormatted
          ),

        recipient_count:
          String(
            totalAttendees
          ),

        external:
          String(
            hasExternalAttendees
          ),

        external_line:
          hasExternalAttendees
            ? '**External Email:** Yes'
            : '',

        body:
          body.trim() ||
          '',
      }
    );

  const safeTitle =
    title.replace(
      /[<>:"/\\|?*]/g,
      '-'
    );

  return {
    kind:
      'email',

    title,

    time,

    fromName:
      fromLine,

    fromEmail,

    location:
      '',

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

function getItemProperty(
  propOrValue
) {
  return new Promise(
    resolve => {
      if (
        propOrValue ===
          undefined ||
        propOrValue ===
          null
      ) {
        resolve('');
        return;
      }

      if (
        typeof propOrValue !==
          'object' ||
        propOrValue instanceof
          Date
      ) {
        resolve(
          propOrValue
        );

        return;
      }

      if (
        typeof propOrValue
          .getAsync ===
        'function'
      ) {
        propOrValue
          .getAsync(
            result => {
              resolve(
                result.status ===
                  Office.AsyncResultStatus.Succeeded
                  ? result.value
                  : ''
              );
            }
          );

      } else {
        resolve('');
      }
    }
  );
}


function getBodyText(item) {
  return new Promise(
    resolve => {
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
              ? (
                  result.value ||
                  ''
                )
              : ''
          );
        }
      );
    }
  );
}


/**
 * Get the body as Markdown.
 *
 * Pipeline:
 *
 * Outlook HTML
 *   ↓
 * cleanOutlookHtml()
 *   ↓
 * removeTeamsInviteBlockHtml()
 *   ↓
 * normalizeOutlookListNesting()
 *   ↓
 * Turndown
 *   ↓
 * removeTeamsInviteText()
 *   ↓
 * compactMarkdown()
 */
function getBodyMarkdown(item) {
  return new Promise(
    resolve => {
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
              .then(
                text => {
                  resolve(
                    compactMarkdown(
                      removeTeamsInviteText(
                        text
                      )
                    )
                  );
                }
              );

            return;
          }

          const html =
            flattenTableCellBlocks(
              promoteOutlookTableHeaders(
                normalizeOutlookListNesting(
                  removeTeamsInviteBlockHtml(
                    cleanOutlookHtml(
                      result.value
                    )
                  )
                )
              )
            );

          if (
            typeof TurndownService ===
            'undefined'
          ) {
            const tmp =
              document
                .createElement(
                  'div'
                );

            tmp.innerHTML =
              html;

            resolve(
              compactMarkdown(
                removeTeamsInviteText(
                  tmp.textContent ||
                  ''
                )
              )
            );

            return;
          }

          try {
            const td =
              new TurndownService(
                {
                  headingStyle:
                    'atx',

                  bulletListMarker:
                    '-',

                  codeBlockStyle:
                    'fenced',

                  emDelimiter:
                    '*',
                }
              );

            if (
              typeof turndownPluginGfm !==
              'undefined'
            ) {
              td.use(
                turndownPluginGfm.gfm
              );
            }

            td.addRule(
              'strip-cid-images',
              {
                filter:
                  node =>
                    node.nodeName ===
                      'IMG' &&
                    (
                      node.getAttribute(
                        'src'
                      ) ||
                      ''
                    )
                      .toLowerCase()
                      .startsWith(
                        'cid:'
                      ),

                replacement:
                  (
                    _content,
                    node
                  ) => {
                    const alt =
                      node.getAttribute(
                        'alt'
                      );

                    return alt
                      ? `*[${alt}]*`
                      : '';
                  },
              }
            );

            const md =
              td
                .turndown(
                  html
                )
                .replace(
                  /\u00a0/g,
                  ' '
                )
                .replace(
                  /\n{3,}/g,
                  '\n\n'
                )
                .trim();

            resolve(
              compactMarkdown(
                removeTeamsInviteText(
                  md
                )
              )
            );

          } catch (e) {
            console.warn(
              'Turndown conversion failed, using text fallback:',
              e.message
            );

            getBodyText(item)
              .then(
                text => {
                  resolve(
                    compactMarkdown(
                      removeTeamsInviteText(
                        text
                      )
                    )
                  );
                }
              );
          }
        }
      );
    }
  );
}


/* ═══════════════════════════════════════════════════════════════════════════
   TEAMS BLOCK REMOVAL — HTML
   ═══════════════════════════════════════════════════════════════════════════ */

function removeTeamsInviteBlockHtml(
  html
) {
  const wrapper =
    document
      .createElement(
        'div'
      );

  wrapper.innerHTML =
    html ||
    '';

  const knownTeamsSelectors = [
    '[id*="MicrosoftTeamsMeetingPlaceholder"]',
    '[id*="TeamsMeetingPlaceholder"]',
    '[class*="MicrosoftTeamsMeetingPlaceholder"]',
    '[class*="TeamsMeetingPlaceholder"]',
  ];

  for (
    const selector
    of knownTeamsSelectors
  ) {
    for (
      const node
      of wrapper
        .querySelectorAll(
          selector
        )
    ) {
      node.remove();
    }
  }

  const blockTags =
    new Set(
      [
        'DIV',
        'P',
        'TABLE',
        'TR',
        'TD',
        'LI',
      ]
    );

  const normalizeText =
    value =>
      String(
        value ||
        ''
      )
        .replace(
          /\u00a0/g,
          ' '
        )
        .replace(
          /\s+/g,
          ' '
        )
        .trim();

  const isTeamsHeading =
    value =>
      /^(?:Microsoft Teams(?:-möte| meeting))$/i
        .test(
          normalizeText(
            value
          )
        );

  const isSeparator =
    value =>
      /^[_\-—–]{20,}$/
        .test(
          String(
            value ||
            ''
          )
            .replace(
              /\u00a0/g,
              ' '
            )
            .replace(
              /\s+/g,
              ''
            )
            .trim()
        );

  const isOrganizerLine =
    value =>
      /^(?:För organisatörer:|For organizers:)/i
        .test(
          normalizeText(
            value
          )
        );

  const allElements =
    Array.from(
      wrapper
        .querySelectorAll(
          '*'
        )
    );

  for (
    const element
    of allElements
  ) {
    if (
      !element.isConnected
    ) {
      continue;
    }

    const elementText =
      normalizeText(
        element.textContent ||
        ''
      );

    if (
      !isTeamsHeading(
        elementText
      )
    ) {
      continue;
    }

    let startNode =
      element;

    while (
      startNode.parentElement &&
      startNode.parentElement !==
        wrapper &&
      !blockTags.has(
        startNode.nodeName
      )
    ) {
      startNode =
        startNode
          .parentElement;
    }

    let node =
      startNode;

    let organizerSeen =
      false;

    let removedCount =
      0;

    while (
      node &&
      node !== wrapper
    ) {
      const next =
        node
          .nextElementSibling;

      const nodeText =
        normalizeText(
          node.textContent ||
          ''
        );

      if (
        isOrganizerLine(
          nodeText
        )
      ) {
        organizerSeen =
          true;
      }

      const separator =
        isSeparator(
          nodeText
        );

      node.remove();

      removedCount +=
        1;

      if (
        separator &&
        removedCount > 1
      ) {
        break;
      }

      if (
        organizerSeen &&
        next
      ) {
        const nextText =
          normalizeText(
            next.textContent ||
            ''
          );

        if (
          !isSeparator(
            nextText
          ) &&
          !/^(?:Privacy and security|Sekretess och säkerhet)/i
            .test(
              nextText
            )
        ) {
          break;
        }
      }

      node =
        next;
    }
  }

  for (
    const link
    of wrapper
      .querySelectorAll(
        'a'
      )
  ) {
    const href =
      (
        link.getAttribute(
          'href'
        ) ||
        ''
      )
        .toLowerCase();

    const linkText =
      normalizeText(
        link.textContent ||
        ''
      )
        .toLowerCase();

    if (
      href.includes(
        'teams.microsoft.com'
      ) ||
      href.includes(
        'teams.live.com'
      ) ||
      href.includes(
        'aka.ms/jointeamsmeeting'
      ) ||
      linkText.includes(
        'join microsoft teams'
      ) ||
      linkText.includes(
        'join teams meeting'
      ) ||
      linkText ===
        'behöver du hjälp?' ||
      linkText ===
        'need help?'
    ) {
      const container =
        link.parentElement;

      const onlyContent =
        container &&
        normalizeText(
          container.textContent
        ) ===
        normalizeText(
          link.textContent
        );

      if (
        onlyContent &&
        [
          'P',
          'DIV',
          'LI',
        ]
          .includes(
            container.nodeName
          )
      ) {
        container.remove();

      } else {
        link.remove();
      }
    }
  }

  return wrapper.innerHTML;
}


/* ═══════════════════════════════════════════════════════════════════════════
   TEAMS BLOCK REMOVAL — TEXT FALLBACK
   ═══════════════════════════════════════════════════════════════════════════ */

function removeTeamsInviteText(
  text
) {
  let result =
    String(
      text ||
      ''
    )
      .replace(
        /\r\n/g,
        '\n'
      );

  result =
    result.replace(
      /(?:^|\n)[ \t]*(?:[_\-—–]{10,})?[ \t]*\n?[ \t]*(?:\*\*)?Microsoft Teams(?:-möte| meeting)(?:\*\*)?[ \t]*\n[\s\S]*?(?=\n[ \t]*(?:[_\-—–]{10,})[ \t]*(?:\n|$)|$)/gi,
      '\n'
    );

  result =
    result.replace(
      /(?:^|\n)[ \t]*(?:\*\*)?Microsoft Teams(?:-möte| meeting)(?:\*\*)?[ \t]*\n[\s\S]*?(?:För organisatörer:|For organizers:)[^\n]*(?:\n|$)/gi,
      '\n'
    );

  result =
    result.replace(
      /https?:\/\/(?:[\w-]+\.)?(?:teams\.microsoft\.com|teams\.live\.com)\/\S+/gi,
      ''
    );

  result =
    result.replace(
      /https?:\/\/aka\.ms\/JoinTeamsMeeting\S*/gi,
      ''
    );

  result =
    result.replace(
      /(?:^|\n)[ \t]*(?:\[[^\]]*\]\([^)]+\)|(?:Behöver du hjälp\?|Need help\?))[ \t]*(?:\|)?[ \t]*(?=\n|$)/gi,
      '\n'
    );

  return result
    .replace(
      /[ \t]+\n/g,
      '\n'
    )
    .replace(
      /\n{3,}/g,
      '\n\n'
    )
    .trim();
}


/* ═══════════════════════════════════════════════════════════════════════════
   OUTLOOK LIST NORMALIZATION
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Repairs list nesting when Outlook presents a child <ul>/<ol> as a sibling
 * rather than as a child of the previous <li>.
 */
function normalizeOutlookListNesting(
  html
) {
  const wrapper =
    document
      .createElement(
        'div'
      );

  wrapper.innerHTML =
    html ||
    '';


  function getMsoListInfo(
    element
  ) {
    const candidates = [
      element,
      ...element
        .querySelectorAll(
          '[style*="mso-list"]'
        ),
    ];

    for (
      const candidate
      of candidates
    ) {
      const style =
        candidate
          .getAttribute(
            'style'
          ) ||
        '';

      const match =
        style.match(
          /mso-list:\s*([^;]*?)\blevel(\d+)\b([^;]*)/i
        );

      if (match) {
        const listIdMatch =
          `${match[1]} ${match[3]}`
            .match(
              /\bl(\d+)\b/i
            );

        const lfoMatch =
          `${match[1]} ${match[3]}`
            .match(
              /\blfo(\d+)\b/i
            );

        return {
          level:
            Number(
              match[2]
            ),

          listId:
            listIdMatch
              ? listIdMatch[1]
              : null,

          lfo:
            lfoMatch
              ? lfoMatch[1]
              : null,
        };
      }
    }

    return {
      level:
        null,

      listId:
        null,

      lfo:
        null,
    };
  }


  function cssLengthToPx(
    value,
    unit
  ) {
    const n =
      Number(
        value
      );

    if (
      !Number.isFinite(
        n
      )
    ) {
      return null;
    }

    switch (
      (
        unit ||
        'px'
      )
        .toLowerCase()
    ) {
      case 'pt':
        return n *
          (96 / 72);

      case 'cm':
        return n *
          (96 / 2.54);

      case 'mm':
        return n *
          (96 / 25.4);

      case 'in':
        return n *
          96;

      default:
        return n;
    }
  }


  function getIndent(
    element
  ) {
    const candidates = [
      element,
      element
        .querySelector(
          'li'
        ),
      element
        .querySelector(
          '[style*="margin-left"]'
        ),
      element
        .querySelector(
          '[style*="text-indent"]'
        ),
    ]
      .filter(
        Boolean
      );

    for (
      const candidate
      of candidates
    ) {
      const style =
        candidate
          .getAttribute(
            'style'
          ) ||
        '';

      const marginMatch =
        style.match(
          /margin-left\s*:\s*(-?\d+(?:\.\d+)?)\s*(px|pt|cm|mm|in)?/i
        );

      if (
        marginMatch
      ) {
        const px =
          cssLengthToPx(
            marginMatch[1],
            marginMatch[2]
          );

        if (
          px !== null
        ) {
          return px;
        }
      }
    }

    return null;
  }


  function isEffectivelyEmpty(
    node
  ) {
    if (!node) {
      return true;
    }

    if (
      node.nodeType ===
      Node.TEXT_NODE
    ) {
      return !(
        node.textContent ||
        ''
      )
        .replace(
          /\u00a0/g,
          ' '
        )
        .trim();
    }

    if (
      node.nodeType !==
      Node.ELEMENT_NODE
    ) {
      return true;
    }

    const text =
      (
        node.textContent ||
        ''
      )
        .replace(
          /\u00a0/g,
          ' '
        )
        .trim();

    if (text) {
      return false;
    }

    return !node
      .querySelector(
        'img, table, hr, br:not(:only-child)'
      );
  }


  function areEffectivelyAdjacent(
    first,
    second
  ) {
    let node =
      first.nextSibling;

    while (
      node &&
      node !== second
    ) {
      if (
        !isEffectivelyEmpty(
          node
        )
      ) {
        return false;
      }

      node =
        node.nextSibling;
    }

    return node === second;
  }


  function sameOutlookList(
    previousInfo,
    currentInfo
  ) {
    if (
      previousInfo.listId !==
        null &&
      currentInfo.listId !==
        null &&
      previousInfo.listId !==
        currentInfo.listId
    ) {
      return false;
    }

    if (
      previousInfo.lfo !==
        null &&
      currentInfo.lfo !==
        null &&
      previousInfo.lfo !==
        currentInfo.lfo
    ) {
      return false;
    }

    return true;
  }


  let changed =
    true;

  let safetyCounter =
    0;

  while (
    changed &&
    safetyCounter <
      1000
  ) {
    changed =
      false;

    safetyCounter +=
      1;

    const lists =
      Array.from(
        wrapper
          .querySelectorAll(
            'ol, ul'
          )
      );

    for (
      const currentList
      of lists
    ) {
      if (
        !currentList.isConnected
      ) {
        continue;
      }

      let previous =
        currentList
          .previousElementSibling;

      while (
        previous &&
        isEffectivelyEmpty(
          previous
        )
      ) {
        previous =
          previous
            .previousElementSibling;
      }

      if (
        !previous ||
        ![
          'OL',
          'UL',
        ]
          .includes(
            previous.nodeName
          )
      ) {
        continue;
      }

      if (
        !areEffectivelyAdjacent(
          previous,
          currentList
        )
      ) {
        continue;
      }

      const previousInfo =
        getMsoListInfo(
          previous
        );

      const currentInfo =
        getMsoListInfo(
          currentList
        );

      const previousIndent =
        getIndent(
          previous
        );

      const currentIndent =
        getIndent(
          currentList
        );

      const nestedByMsoLevel =
        previousInfo.level !==
          null &&
        currentInfo.level !==
          null &&
        sameOutlookList(
          previousInfo,
          currentInfo
        ) &&
        currentInfo.level >
          previousInfo.level;

      const nestedByIndent =
        previousIndent !==
          null &&
        currentIndent !==
          null &&
        currentIndent >
          previousIndent +
          8;

      const noContradictingLevel =
        previousInfo.level ===
          null ||
        currentInfo.level ===
          null ||
        currentInfo.level >
          previousInfo.level;

      const nestedByCommonPattern =
        previous.nodeName ===
          'OL' &&
        currentList.nodeName ===
          'UL' &&
        noContradictingLevel;

      if (
        !nestedByMsoLevel &&
        !nestedByIndent &&
        !nestedByCommonPattern
      ) {
        continue;
      }

      const lastItem =
        Array.from(
          previous.children
        )
          .reverse()
          .find(
            child =>
              child.nodeName ===
              'LI'
          );

      if (
        !lastItem
      ) {
        continue;
      }

      lastItem
        .appendChild(
          currentList
        );

      changed =
        true;

      break;
    }
  }

  return wrapper.innerHTML;
}


/* ═══════════════════════════════════════════════════════════════════════════
   COMPACT MARKDOWN
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Final cleanup after Turndown.
 *
 * Keeps paragraphs readable while making lists compact.
 */
function compactMarkdown(
  markdown
) {
  let result =
    String(
      markdown ||
      ''
    )
      .replace(
        /\r\n?/g,
        '\n'
      )
      .replace(
        /\u00a0/g,
        ' '
      )
      .replace(
        /[ \t]+$/gm,
        ''
      );


  /*
   * Remove long Outlook separator lines.
   *
   * Normal Markdown "---" is left intact.
   */
  result =
    result
      .replace(
        /^[ \t]*_{20,}[ \t]*$/gm,
        ''
      )
      .replace(
        /^[ \t]*-{10,}[ \t]*$/gm,
        ''
      );


  /*
   * Remove empty list items.
   *
   * Examples:
   *
   *   5.
   *   5)
   *   -
   */
  result =
    result
      .replace(
        /^[ \t]*\d+[.)][ \t]*$/gm,
        ''
      )
      .replace(
        /^[ \t]*[-*+][ \t]*$/gm,
        ''
      );


  /*
   * Remove blank lines between consecutive bullet items.
   */
  result =
    result.replace(
      /^([ \t]*[-*+] .+)\n[ \t]*\n(?=[ \t]*[-*+] )/gm,
      '$1\n'
    );


  /*
   * Remove blank lines between consecutive numbered items.
   */
  result =
    result.replace(
      /^([ \t]*\d+[.)] .+)\n[ \t]*\n(?=[ \t]*\d+[.)] )/gm,
      '$1\n'
    );


  /*
   * Remove blank line between numbered parent and nested bullet.
   */
  result =
    result.replace(
      /^([ \t]*\d+[.)] .+)\n[ \t]*\n(?=[ \t]+[-*+] )/gm,
      '$1\n'
    );


  /*
   * Remove blank line between nested bullet and next numbered item.
   */
  result =
    result.replace(
      /^([ \t]+[-*+] .+)\n[ \t]*\n(?=[ \t]*\d+[.)] )/gm,
      '$1\n'
    );


  /*
   * Second compaction pass.
   */
  result =
    result
      .replace(
        /^([ \t]*[-*+] .+)\n[ \t]*\n(?=[ \t]*[-*+] )/gm,
        '$1\n'
      )
      .replace(
        /^([ \t]*\d+[.)] .+)\n[ \t]*\n(?=[ \t]*\d+[.)] )/gm,
        '$1\n'
      )
      .replace(
        /^([ \t]*\d+[.)] .+)\n[ \t]*\n(?=[ \t]+[-*+] )/gm,
        '$1\n'
      )
      .replace(
        /^([ \t]+[-*+] .+)\n[ \t]*\n(?=[ \t]*\d+[.)] )/gm,
        '$1\n'
      );


  /*
   * Never retain more than one empty line.
   */
  return result
    .replace(
      /\n{3,}/g,
      '\n\n'
    )
    .trim();
}


/* ═══════════════════════════════════════════════════════════════════════════
   TABLE NORMALIZATION
   ═══════════════════════════════════════════════════════════════════════════ */

function promoteOutlookTableHeaders(
  html
) {
  const wrapper =
    document
      .createElement(
        'div'
      );

  wrapper.innerHTML =
    html;

  for (
    const table
    of wrapper
      .querySelectorAll(
        'table'
      )
  ) {
    if (
      table.querySelector(
        'th'
      ) ||
      table.querySelector(
        'thead'
      )
    ) {
      continue;
    }

    const firstRow =
      table
        .querySelector(
          'tr'
        );

    if (
      !firstRow
    ) {
      continue;
    }

    const tdCells =
      Array.from(
        firstRow.children
      )
        .filter(
          cell =>
            cell.nodeName ===
            'TD'
        );

    if (
      tdCells.length ===
      0
    ) {
      continue;
    }

    for (
      const td
      of tdCells
    ) {
      const th =
        document
          .createElement(
            'th'
          );

      for (
        const attr
        of Array.from(
          td.attributes
        )
      ) {
        th.setAttribute(
          attr.name,
          attr.value
        );
      }

      while (
        td.firstChild
      ) {
        th.appendChild(
          td.firstChild
        );
      }

      td.parentNode
        .replaceChild(
          th,
          td
        );
    }
  }

  return wrapper.innerHTML;
}


function flattenTableCellBlocks(
  html
) {
  const wrapper =
    document
      .createElement(
        'div'
      );

  wrapper.innerHTML =
    html;

  for (
    const cell
    of wrapper
      .querySelectorAll(
        'td, th'
      )
  ) {
    let block;

    while (
      (
        block =
          cell.querySelector(
            'p, div'
          )
      )
    ) {
      const parent =
        block.parentNode;

      while (
        block.firstChild
      ) {
        parent.insertBefore(
          block.firstChild,
          block
        );
      }

      parent.removeChild(
        block
      );
    }
  }

  return wrapper.innerHTML;
}


/* ═══════════════════════════════════════════════════════════════════════════
   OUTLOOK HTML CLEANUP
   ═══════════════════════════════════════════════════════════════════════════ */

function cleanOutlookHtml(
  html
) {
  return html

    /*
     * MSO conditional comments.
     */
    .replace(
      /<!--\[if[^>]*?\]>[\s\S]*?<!\[endif\]-->/gi,
      ''
    )

    /*
     * Outlook <o:p> elements.
     */
    .replace(
      /<o:p[^>]*>[\s\S]*?<\/o:p>/gi,
      ''
    )
    .replace(
      /<o:p[^>]*\/>/gi,
      ''
    )

    /*
     * Style blocks.
     */
    .replace(
      /<style[\s\S]*?<\/style>/gi,
      ''
    )

    /*
     * Remaining metadata/head leftovers.
     */
    .replace(
      /<\/?(?:meta|link|style)\b[^>]*>/gi,
      ''
    );
}


/* ═══════════════════════════════════════════════════════════════════════════
   ATTENDEE EXTRACTION — OFFICE.JS
   ═══════════════════════════════════════════════════════════════════════════ */

function getAttendeesViaOfficeJs(
  item
) {
  return new Promise(
    async resolve => {
      const result = {
        'Accepted':
          new Map(),

        'Tentative':
          new Map(),

        'Declined':
          new Map(),

        'No Response':
          new Map(),
      };

      const RESPONSE_MAP = {
        accepted:
          'Accepted',

        tentative:
          'Tentative',

        declined:
          'Declined',

        none:
          'No Response',

        organizer:
          null,
      };


      async function fetchList(
        prop
      ) {
        const value =
          item[prop];

        if (!value) {
          return [];
        }

        if (
          Array.isArray(
            value
          )
        ) {
          return value;
        }

        if (
          typeof value
            .getAsync ===
          'function'
        ) {
          return await new Promise(
            res => {
              value.getAsync(
                r => {
                  res(
                    r.status ===
                      Office.AsyncResultStatus.Succeeded
                      ? (
                          r.value ||
                          []
                        )
                      : []
                  );
                }
              );
            }
          );
        }

        return [];
      }


      const required =
        await fetchList(
          'requiredAttendees'
        );

      const optional =
        await fetchList(
          'optionalAttendees'
        );

      for (
        const attendee
        of [
          ...required,
          ...optional,
        ]
      ) {
        const name =
          attendee.displayName ||
          attendee.emailAddress ||
          '';

        const email =
          (
            attendee.emailAddress ||
            ''
          )
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
          Object.prototype
            .hasOwnProperty
            .call(
              RESPONSE_MAP,
              responseType
            )
            ? RESPONSE_MAP[
                responseType
              ]
            : 'No Response';

        if (
          bucket === null
        ) {
          continue;
        }

        result[
          bucket
        ].set(
          name,
          email
        );
      }

      resolve(
        result
      );
    }
  );
}


/* ═══════════════════════════════════════════════════════════════════════════
   EWS — ATTENDEES WITH RSVP STATUS
   ═══════════════════════════════════════════════════════════════════════════ */

function getAttendeesViaEws(
  itemId
) {
  return new Promise(
    (
      resolve,
      reject
    ) => {
      let ewsId =
        itemId;

      try {
        const converted =
          Office.context
            .mailbox
            .convertToEwsId(
              itemId,
              Office.MailboxEnums
                .RestVersion
                .v2_0
            );

        if (
          converted
        ) {
          ewsId =
            converted;
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

      Office.context
        .mailbox
        .makeEwsRequestAsync(
          soap,
          result => {
            if (
              result.status !==
              Office.AsyncResultStatus.Succeeded
            ) {
              const err =
                result.error ||
                {};

              const code =
                err.code != null
                  ? ` code=${err.code}`
                  : '';

              const name =
                err.name
                  ? ` name=${err.name}`
                  : '';

              const body =
                typeof result.value ===
                  'string' &&
                result.value.length > 0
                  ? (
                      ' body=' +
                      result.value
                        .replace(
                          /\s+/g,
                          ' '
                        )
                        .slice(
                          0,
                          400
                        )
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
    }
  );
}


/* ═══════════════════════════════════════════════════════════════════════════
   EWS PARSING
   ═══════════════════════════════════════════════════════════════════════════ */

function parseEwsAttendees(
  xmlString
) {
  const parser =
    new DOMParser();

  const doc =
    parser
      .parseFromString(
        xmlString,
        'text/xml'
      );

  const T =
    'http://schemas.microsoft.com/exchange/services/2006/types';

  const M =
    'http://schemas.microsoft.com/exchange/services/2006/messages';


  /* ── Surface EWS errors ───────────────────────────────────────────────── */

  const responseMsgs =
    doc
      .getElementsByTagNameNS(
        M,
        'GetItemResponseMessage'
      );

  if (
    responseMsgs.length >
    0
  ) {
    const responseClass =
      responseMsgs[0]
        .getAttribute(
          'ResponseClass'
        );

    if (
      responseClass ===
      'Error'
    ) {
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
    'Accepted':
      new Map(),

    'Tentative':
      new Map(),

    'Declined':
      new Map(),

    'No Response':
      new Map(),
  };

  const RSVP_MAP = {
    accept:
      'Accepted',

    tentative:
      'Tentative',

    decline:
      'Declined',

    none:
      'No Response',

    noresponsereceived:
      'No Response',

    organizer:
      null,
  };


  function processAttendeeGroup(
    tagName
  ) {
    for (
      const group
      of doc
        .getElementsByTagNameNS(
          T,
          tagName
        )
    ) {
      for (
        const attendeeEl
        of group
          .getElementsByTagNameNS(
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

        if (
          !name
        ) {
          continue;
        }

        const bucket =
          Object.prototype
            .hasOwnProperty
            .call(
              RSVP_MAP,
              responseType
            )
            ? RSVP_MAP[
                responseType
              ]
            : 'No Response';

        if (
          bucket === null
        ) {
          continue;
        }

        result[
          bucket
        ].set(
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