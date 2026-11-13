
/**
 * TASK HUB | RECURRING TASK AUTOMATION
 *
 * Complete Google Apps Script replacement.
 * Workbook layout: Cycle Tracker columns A:U (21 columns).
 *
 * Employee updates: Employee View
 * Manager acknowledgements: Manager Dashboard
 * History: Completion Log
 * Notifications: Alert Log
 *
 * Dry Run Mode remains controlled by SLA & Settings.
 * Edit submissions read the actual edited cell value (range.getValue).
 * Display locale can be set to English (India) from the Task Hub menu.
 */

const CFG = {
  SHEETS: {
    TASKS: 'Task Master',
    TEAM: 'Team',
    CYCLES: 'Cycle Tracker',
    EMPLOYEE: 'Employee View',
    MANAGER: 'Manager Dashboard',
    COMPLETIONS: 'Completion Log',
    ALERTS: 'Alert Log',
    SETTINGS: 'SLA & Settings'
  },

  HEADER_ROW: 4,
  DATA_ROW: 5,
  MANAGER_HEADER_ROW: 12,
  MANAGER_DATA_ROW: 13,

  // Cycle Tracker: exactly 21 columns, A through U.
  CYCLE: {
    ID: 1,
    TASK_ID: 2,
    TASK: 3,
    OWNER: 4,
    NUMBER: 5,
    START: 6,
    DUE: 7,
    DUE_TIME: 8,
    STATUS: 9,
    REQUIRED_PROOF: 10,
    PROOF: 11,
    PROOF_LINK: 12,
    COMPLETED_BY: 13,
    COMPLETED_AT: 14,
    MANAGER_ALERT_AT: 15,
    MANAGER_ACK_BY: 16,
    MANAGER_ACK_AT: 17,
    LEADERSHIP_AT: 18,
    UPDATED_AT: 19,
    NOTES: 20,
    IDENTITY_SOURCE: 21
  },

  // Employee View: A:J.
  EMPLOYEE: {
    ID: 1,
    TASK: 2,
    OWNER: 3,
    DUE: 4,
    DUE_TIME: 5,
    DUE_STATE: 6,
    REQUIRED_PROOF: 7,
    PROOF: 8,
    PROOF_LINK: 9,
    SUBMIT: 10
  },

  // Manager Dashboard: A:N.
  MANAGER: {
    ID: 1,
    OWNER: 2,
    TASK: 3,
    DUE: 4,
    DUE_TIME: 5,
    STATUS: 6,
    DUE_STATE: 7,
    ALERT_AT: 8,
    REVIEWED: 9,
    REVIEWED_BY: 10,
    REVIEWED_AT: 11,
    LEADERSHIP_AT: 12,
    REQUIRED_PROOF: 13,
    PROOF: 14
  }
};


/* =========================================================
   1. MENU AND WORKBOOK VALIDATION
   ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Task Hub')
    .addItem('Refresh views', 'refreshViews')
    .addItem('Run automation now', 'runAutomation')
    .addItem('Set date display to English (India)', 'setEnglishIndiaDisplay')
    .addSeparator()
    .addItem('Set up automation', 'setupAutomation')
    .addToUi();
}


/** Changes spreadsheet locale only; does not alter cell values or triggers. */
function setEnglishIndiaDisplay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('en_IN');
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  SpreadsheetApp.flush();
  ss.toast('Date and number display is now set to English (India).', 'Task Hub', 7);
}


/**
 * Checks the actual workbook headers before any data is read
 * or written. This helps prevent column mismatches.
 */
function validateWorkbookStructure_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const checks = [
    {
      sheet: CFG.SHEETS.CYCLES,
      row: 4,
      headers: [
        'Cycle ID',
        'Task ID',
        'Task Name',
        'Assigned To',
        'Cycle #',
        'Cycle Start Date',
        'Due Date',
        'Due Time',
        'Status',
        'Required Proof',
        'Ticket ID / Completion Note',
        'Proof Link',
        'Completed By',
        'Completed At',
        'Manager Alert Sent At',
        'Manager Reviewed By',
        'Manager Reviewed At',
        'Leadership Escalated At',
        'Last Updated At',
        'Notes',
        'Identity Source'
      ]
    },

    {
      sheet: CFG.SHEETS.EMPLOYEE,
      row: 4,
      headers: [
        'Cycle ID',
        'Task',
        'Assigned To',
        'Due Date',
        'Due Time',
        'Due status',
        'What to provide',
        'Ticket ID / Completion note',
        'Proof link (optional)',
        'Submit completion?'
      ]
    },

    {
      sheet: CFG.SHEETS.MANAGER,
      row: 12,
      headers: [
        'Cycle ID',
        'Employee',
        'Task',
        'Due Date',
        'Due Time',
        'Task Status',
        'Due State',
        'Manager Alert Sent At',
        'Manager reviewed?',
        'Reviewed By',
        'Reviewed At',
        'Leadership Escalated At',
        'Required Proof',
        'Ticket ID / Completion note'
      ]
    },

    {
      sheet: CFG.SHEETS.TASKS,
      row: 4,
      headers: [
        'Task ID',
        'Task Name',
        'Assigned To',
        'Frequency',
        'Interval (days)',
        'Weekdays only?',
        'Due time',
        'Task type',
        'Required proof',
        'Active?',
        'First due date',
        'Notes'
      ]
    }
  ];

  checks.forEach(function (check) {

    const sheet = ss.getSheetByName(check.sheet);

    if (!sheet) {
      throw new Error('Required tab not found: ' + check.sheet);
    }

    const actual = sheet
      .getRange(check.row, 1, 1, check.headers.length)
      .getDisplayValues()[0]
      .map(function (value) {
        return String(value || '').trim();
      });

    const mismatches = [];

    check.headers.forEach(function (expected, index) {

      if (actual[index] !== expected) {

        mismatches.push(
          'Column ' + (index + 1) +
          ': expected "' + expected +
          '", found "' + actual[index] + '"'
        );

      }

    });

    if (mismatches.length) {

      throw new Error(
        'The headers in "' + check.sheet +
        '" do not match the approved workbook layout. ' +
        mismatches.join('; ')
      );

    }

  });

}


/**
 * RECOVERY FUNCTION
 *
 * The earlier troubleshooting function accidentally moved
 * manager alert timestamps from O into P.
 *
 * This function moves them back ONLY when:
 * 1. The cycle is pending.
 * 2. Column O is blank.
 * 3. Column P contains a real date/time.
 * 4. Column Q is blank.
 *
 * It does not alter other columns or completion records.
 *
 * It runs automatically when refreshing or running automation.
 */
function repairLegacyShiftedAlertTimestamps_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.CYCLES);

  const lastRow = sheet.getLastRow();

  if (lastRow < CFG.DATA_ROW) return 0;

  const rows = sheet.getRange(
    CFG.DATA_ROW,
    1,
    lastRow - CFG.DATA_ROW + 1,
    21
  ).getValues();

  let repaired = 0;

  rows.forEach(function (row, index) {

    const status = String(
      row[CFG.CYCLE.STATUS - 1] || ''
    ).trim().toLowerCase();

    const managerAlertCell =
      row[CFG.CYCLE.MANAGER_ALERT_AT - 1]; // O

    const managerReviewedByCell =
      row[CFG.CYCLE.MANAGER_ACK_BY - 1]; // P

    const managerReviewedAtCell =
      row[CFG.CYCLE.MANAGER_ACK_AT - 1]; // Q

    const alertCellIsBlank =
      managerAlertCell === '' ||
      managerAlertCell === null;

    const looksLikeShiftedAlert =
      isDate_(managerReviewedByCell);

    const reviewTimestampIsBlank =
      managerReviewedAtCell === '' ||
      managerReviewedAtCell === null;

    if (
      status === 'pending' &&
      alertCellIsBlank &&
      looksLikeShiftedAlert &&
      reviewTimestampIsBlank
    ) {

      const sheetRow = CFG.DATA_ROW + index;

      // Restore the timestamp to column O.
      sheet.getRange(
        sheetRow,
        CFG.CYCLE.MANAGER_ALERT_AT
      ).setValue(managerReviewedByCell);

      // Clear the misplaced value from column P.
      sheet.getRange(
        sheetRow,
        CFG.CYCLE.MANAGER_ACK_BY
      ).clearContent();

      repaired++;

    }

  });

  return repaired;

}


/* =========================================================
   2. AUTOMATION SETUP AND REFRESH
   ========================================================= */


/**
 * Installs the two required triggers.
 * Running this again will replace existing triggers rather
 * than creating duplicates.
 */
function setupAutomation() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Open this script from the Task Hub Google Sheet.'
    );
  }

  validateWorkbookStructure_();

  ss.setSpreadsheetTimeZone('Asia/Kolkata');

  const repaired = repairLegacyShiftedAlertTimestamps_();

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (trigger) {

    const fn = trigger.getHandlerFunction();

    if (
      fn === 'runAutomation' ||
      fn === 'handleEdit'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

  ScriptApp
    .newTrigger('runAutomation')
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp
    .newTrigger('handleEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ensureInitialCycles_();

  refreshViews_();

  ss.toast(
    'Automation is set up. ' +
    repaired +
    ' earlier alert timestamp(s) restored.',
    'Task Hub',
    8
  );

}


/**
 * Refreshes Employee View and Manager Dashboard.
 * Also restores the two timestamps shifted during troubleshooting.
 */
function refreshViews() {

  validateWorkbookStructure_();

  const repaired =
    repairLegacyShiftedAlertTimestamps_();

  ensureInitialCycles_();

  refreshViews_();

  const message = repaired
    ? 'Views refreshed. ' + repaired +
      ' earlier alert timestamp(s) restored to the correct column.'
    : 'The employee and manager views are up to date.';

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(message, 'Task Hub', 8);

}


function refreshViews_() {

  refreshEmployeeView_();

  refreshManagerDashboard_();

}


/* =========================================================
   3. MAIN AUTOMATION
   ========================================================= */


/**
 * Runs every 15 minutes.
 *
 * Checks overdue tasks, sends manager alerts, checks the
 * manager acknowledgement SLA and escalates unresolved tasks.
 */
function runAutomation() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(25000)) return;

  try {

    validateWorkbookStructure_();

    repairLegacyShiftedAlertTimestamps_();

    ensureInitialCycles_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName(
      CFG.SHEETS.CYCLES
    );

    const lastRow = sheet.getLastRow();

    if (lastRow < CFG.DATA_ROW) {

      refreshViews_();

      return;

    }

    const values = sheet.getRange(
      CFG.DATA_ROW,
      1,
      lastRow - CFG.DATA_ROW + 1,
      21
    ).getValues();

    const settings = getSettings_();

    const now = new Date();

    values.forEach(function (row, index) {

      const sheetRow = CFG.DATA_ROW + index;

      const cycleId = String(
        row[CFG.CYCLE.ID - 1] || ''
      ).trim();

      const status = String(
        row[CFG.CYCLE.STATUS - 1] || ''
      ).trim().toLowerCase();

      if (!cycleId || status !== 'pending') return;

      const dueAt = combineDateAndTime_(
        row[CFG.CYCLE.DUE - 1],
        row[CFG.CYCLE.DUE_TIME - 1]
      );

      if (!dueAt) return;

      let managerAlertAt =
        row[CFG.CYCLE.MANAGER_ALERT_AT - 1];

      const managerAckAt =
        row[CFG.CYCLE.MANAGER_ACK_AT - 1];

      const leadershipAt =
        row[CFG.CYCLE.LEADERSHIP_AT - 1];


      // STEP 1: Send the first manager alert when a task becomes overdue.

      if (
        now >= dueAt &&
        !isDate_(managerAlertAt)
      ) {

        const result = sendAlert_(
          'Manager',
          row,
          settings,
          now
        );

        if (result.stamp) {

          sheet.getRange(
            sheetRow,
            CFG.CYCLE.MANAGER_ALERT_AT
          ).setValue(now);

          sheet.getRange(
            sheetRow,
            CFG.CYCLE.UPDATED_AT
          ).setValue(now);

          managerAlertAt = now;

        }

        logAlert_(
          cycleId,
          'Manager alert',
          settings.managerEmail || '',
          now,
          result.outcome,
          result.summary
        );

      }


      // STEP 2: Check acknowledgement and resolution deadlines.

      if (
        isDate_(managerAlertAt) &&
        !isDate_(leadershipAt)
      ) {

        const ackHours = numberOr_(
          settings.managerAckHours,
          4
        );

        const resolutionHours = numberOr_(
          settings.resolutionHours,
          24
        );

        const acknowledgementDeadline =
          addBusinessHours_(
            managerAlertAt,
            ackHours,
            settings
          );

        const resolutionDeadline = new Date(
          managerAlertAt.getTime() +
          resolutionHours * 60 * 60 * 1000
        );

        const managerHasNotReviewed =
          !isDate_(managerAckAt);

        const reviewWindowMissed =
          managerHasNotReviewed &&
          now >= acknowledgementDeadline;

        const resolutionWindowMissed =
          now >= resolutionDeadline;


        // STEP 3: Escalate once if either SLA has been missed.

        if (
          reviewWindowMissed ||
          resolutionWindowMissed
        ) {

          const result = sendAlert_(
            'Leadership escalation',
            row,
            settings,
            now
          );

          if (result.stamp) {

            sheet.getRange(
              sheetRow,
              CFG.CYCLE.LEADERSHIP_AT
            ).setValue(now);

            sheet.getRange(
              sheetRow,
              CFG.CYCLE.UPDATED_AT
            ).setValue(now);

          }

          const recipients =
            (settings.leadershipEmails || []).join(', ');

          logAlert_(
            cycleId,
            'Leadership escalation',
            recipients,
            now,
            result.outcome,
            result.summary
          );

        }

      }

    });

    refreshViews_();

  } finally {

    lock.releaseLock();

  }

}


/* =========================================================
   4. EMPLOYEE SUBMISSION AND MANAGER ACKNOWLEDGEMENT
   ========================================================= */


/**
 * Installable edit trigger.
 * Reads the edited cell directly from the range rather than relying
 * only on e.value, which can be absent for some edit/paste paths.
 * Watches for YES in Employee View and Manager Dashboard.
 */
function handleEdit(e) {

  if (!e || !e.range) {
    Logger.log('handleEdit received no range; nothing to process.');
    return;
  }

  const range = e.range;

  // This workflow is intentionally designed for one dropdown cell at a time.
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    Logger.log('Ignored multi-cell edit: ' + range.getA1Notation());
    return;
  }

  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();

  // Use the actual value now in the edited cell. Do not rely on e.value.
  const value = String(range.getValue() || '')
    .trim()
    .toUpperCase();

  Logger.log(JSON.stringify({
    event: 'handleEdit',
    sheet: sheetName,
    cell: range.getA1Notation(),
    row: row,
    column: col,
    cellValue: value
  }));

  if (value !== 'YES') return;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(25000)) {
    range.setValue('No');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Another update is being processed. Please try once more.',
      'Task Hub',
      5
    );
    return;
  }

  try {

    if (
      sheetName === CFG.SHEETS.EMPLOYEE &&
      row >= CFG.DATA_ROW &&
      col === CFG.EMPLOYEE.SUBMIT
    ) {
      Logger.log('Processing employee completion for row ' + row);
      submitEmployeeCompletion_(sheet, row);
      return;
    }

    if (
      sheetName === CFG.SHEETS.MANAGER &&
      row >= CFG.MANAGER_DATA_ROW &&
      col === CFG.MANAGER.REVIEWED
    ) {
      Logger.log('Processing manager acknowledgement for row ' + row);
      recordManagerReview_(sheet, row);
      return;
    }

    Logger.log('YES was selected outside a supported submission cell.');

  } catch (err) {
    Logger.log('handleEdit failed: ' + (err.stack || err.message || err));
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'The update could not be saved: ' + err.message,
      'Task Hub — action needed',
      10
    );
    throw err;

  } finally {
    lock.releaseLock();
  }
}


/**
 * Saves an employee's completion from Employee View.
 * One submission updates the cycle and Completion Log.
 */
function submitEmployeeCompletion_(
  employeeSheet,
  employeeRow
) {

  Logger.log('submitEmployeeCompletion_ entered for Employee View row ' + employeeRow);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cycleId = String(
    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.ID
    ).getValue() || ''
  ).trim();

  const proof = String(
    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.PROOF
    ).getValue() || ''
  ).trim();

  const proofLink = String(
    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.PROOF_LINK
    ).getValue() || ''
  ).trim();


  if (!cycleId) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    return;

  }


  const cycleSheet = ss.getSheetByName(
    CFG.SHEETS.CYCLES
  );

  const cycleRow = findCycleRow_(cycleId);

  if (!cycleRow) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'That cycle could not be found. Refresh the view and try again.',
      'Task Hub',
      6
    );

    return;

  }


  const cycle = cycleSheet.getRange(
    cycleRow,
    1,
    1,
    21
  ).getValues()[0];


  if (
    String(
      cycle[CFG.CYCLE.STATUS - 1]
    ).toLowerCase() !== 'pending'
  ) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'This cycle has already been closed. The view will refresh now.',
      'Task Hub',
      5
    );

    refreshViews_();

    return;

  }


  const taskType = getTaskType_(
    String(cycle[CFG.CYCLE.TASK_ID - 1])
  );


  // Portal tasks must have a ticket ID.

  if (
    taskType === 'portal' &&
    !proof
  ) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'Please add the portal Ticket ID before submitting.',
      'Task Hub',
      6
    );

    return;

  }


  // Daily and other non-portal tasks need a note or proof link.

  if (
    taskType !== 'portal' &&
    !proof &&
    !proofLink
  ) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'Add a short completion note or a proof link before submitting.',
      'Task Hub',
      6
    );

    return;

  }


  const owner = String(
    cycle[CFG.CYCLE.OWNER - 1] || ''
  ).trim();

  const ownerEmail = getTeamEmail_(owner);

  const activeEmail = getActiveEmail_();


  // If a real employee email has been configured, check the submitter.

  if (
    activeEmail &&
    ownerEmail &&
    activeEmail.toLowerCase() !== ownerEmail.toLowerCase()
  ) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'Please submit this task from the assigned employee account.',
      'Task Hub',
      6
    );

    return;

  }


  const completedAt = new Date();

  const completedBy =
    activeEmail ||
    ownerEmail ||
    owner ||
    'Employee';

  const identitySource = activeEmail
    ? 'Google account'
    : (
        ownerEmail
          ? 'Assigned employee email; Google account not exposed'
          : 'Assigned task owner'
      );


  // Prevent a second completion log for the same cycle.

  if (completionAlreadyLogged_(cycleId)) {

    employeeSheet.getRange(
      employeeRow,
      CFG.EMPLOYEE.SUBMIT
    ).setValue('No');

    ss.toast(
      'This cycle has already been recorded as complete.',
      'Task Hub',
      5
    );

    refreshViews_();

    return;

  }


  const taskInfo = getTaskMap_()[
    String(cycle[CFG.CYCLE.TASK_ID - 1])
  ];

  const nextDue = taskInfo
    ? calculateNextDueDate_(completedAt, taskInfo)
    : null;


  // Close the current cycle.

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.STATUS
  ).setValue('Completed');

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.PROOF
  ).setValue(proof);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.PROOF_LINK
  ).setValue(proofLink);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.COMPLETED_BY
  ).setValue(completedBy);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.COMPLETED_AT
  ).setValue(completedAt);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.UPDATED_AT
  ).setValue(completedAt);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.IDENTITY_SOURCE
  ).setValue(identitySource);


  // Add the single completion record.

  appendCompletionLog_(
    cycle,
    completedBy,
    completedAt,
    proof,
    proofLink,
    nextDue,
    identitySource
  );


  // Create the next cycle only if another pending cycle does not exist.

  const taskId = String(
    cycle[CFG.CYCLE.TASK_ID - 1]
  );

  if (
    taskInfo &&
    !hasPendingCycleForTask_(taskId)
  ) {

    const nextCycleNumber =
      getNextCycleNumber_(taskId);

    appendCycle_(
      taskInfo,
      nextCycleNumber,
      completedAt,
      nextDue
    );

  }


  refreshViews_();

  ss.toast(
    'Completion saved. The next cycle has been added to the tracker.',
    'Task Hub',
    6
  );

}


/**
 * Records a manager acknowledgement in the Cycle Tracker.
 */
function recordManagerReview_(
  managerSheet,
  managerRow
) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cycleId = String(
    managerSheet.getRange(
      managerRow,
      CFG.MANAGER.ID
    ).getValue() || ''
  ).trim();

  const cycleRow = findCycleRow_(cycleId);

  const settings = getSettings_();

  const activeEmail = getActiveEmail_();


  if (!cycleRow) {

    managerSheet.getRange(
      managerRow,
      CFG.MANAGER.REVIEWED
    ).setValue('No');

    return;

  }


  const configuredManager = String(
    settings.managerEmail || ''
  ).trim().toLowerCase();


  if (
    activeEmail &&
    configuredManager &&
    activeEmail.toLowerCase() !== configuredManager
  ) {

    managerSheet.getRange(
      managerRow,
      CFG.MANAGER.REVIEWED
    ).setValue('No');

    ss.toast(
      'Please acknowledge this issue from the manager account listed in SLA & Settings.',
      'Task Hub',
      6
    );

    return;

  }


  const cycleSheet = ss.getSheetByName(
    CFG.SHEETS.CYCLES
  );

  const reviewedAt = new Date();

  const reviewedBy =
    activeEmail ||
    settings.managerEmail ||
    'Manager';


  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.MANAGER_ACK_BY
  ).setValue(reviewedBy);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.MANAGER_ACK_AT
  ).setValue(reviewedAt);

  cycleSheet.getRange(
    cycleRow,
    CFG.CYCLE.UPDATED_AT
  ).setValue(reviewedAt);


  refreshViews_();

  ss.toast(
    'Manager review recorded. The employee task is still open until it is completed.',
    'Task Hub',
    6
  );

}


/* =========================================================
   5. TASK AND CYCLE MANAGEMENT
   ========================================================= */


/**
 * Creates an initial cycle for active tasks that do not
 * currently have a pending cycle.
 */
function ensureInitialCycles_() {

  const taskMap = getTaskMap_();

  Object.keys(taskMap).forEach(function (taskId) {

    const task = taskMap[taskId];

    if (
      !task.active ||
      hasPendingCycleForTask_(taskId)
    ) {
      return;
    }

    const firstDue =
      task.firstDueDate ||
      dateOnly_(new Date());

    appendCycle_(
      task,
      getNextCycleNumber_(taskId),
      firstDue,
      firstDue
    );

  });

}


/**
 * Rebuilds Employee View.
 * Any unsubmitted proof notes are preserved by Cycle ID.
 */
function refreshEmployeeView_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(
    CFG.SHEETS.EMPLOYEE
  );

  const oldLast = sheet.getLastRow();

  const preserved = {};


  if (oldLast >= CFG.DATA_ROW) {

    const oldRows = sheet.getRange(
      CFG.DATA_ROW,
      1,
      oldLast - CFG.DATA_ROW + 1,
      10
    ).getValues();

    oldRows.forEach(function (row) {

      const id = String(
        row[0] || ''
      ).trim();

      if (id) {

        preserved[id] = {
          proof: row[7] || '',
          link: row[8] || ''
        };

      }

    });

  }


  const cycles = getCycleRows_().filter(function (row) {

    return String(
      row[CFG.CYCLE.STATUS - 1] || ''
    ).toLowerCase() === 'pending';

  });


  cycles.sort(function (a, b) {

    const da = combineDateAndTime_(
      a[CFG.CYCLE.DUE - 1],
      a[CFG.CYCLE.DUE_TIME - 1]
    );

    const db = combineDateAndTime_(
      b[CFG.CYCLE.DUE - 1],
      b[CFG.CYCLE.DUE_TIME - 1]
    );

    return (
      (da ? da.getTime() : 0) -
      (db ? db.getTime() : 0)
    );

  });


  const now = new Date();


  const output = cycles.map(function (row) {

    const id = String(
      row[CFG.CYCLE.ID - 1] || ''
    );

    const dueAt = combineDateAndTime_(
      row[CFG.CYCLE.DUE - 1],
      row[CFG.CYCLE.DUE_TIME - 1]
    );

    const saved = preserved[id] || {
      proof: '',
      link: ''
    };

    return [
      id,
      row[CFG.CYCLE.TASK - 1],
      row[CFG.CYCLE.OWNER - 1],
      row[CFG.CYCLE.DUE - 1],
      row[CFG.CYCLE.DUE_TIME - 1],
      dueState_(dueAt, now),
      row[CFG.CYCLE.REQUIRED_PROOF - 1],
      saved.proof,
      saved.link,
      'No'
    ];

  });


  if (oldLast >= CFG.DATA_ROW) {

    sheet.getRange(
      CFG.DATA_ROW,
      1,
      oldLast - CFG.DATA_ROW + 1,
      10
    ).clearContent();

  }


  if (output.length) {

    sheet.getRange(
      CFG.DATA_ROW,
      1,
      output.length,
      10
    ).setValues(output);

    sheet.getRange(
      CFG.DATA_ROW,
      4,
      output.length,
      1
    ).setNumberFormat('dd mmm yyyy');

  }

}


/**
 * Rebuilds Manager Dashboard from all currently pending cycles.
 */
function refreshManagerDashboard_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(
    CFG.SHEETS.MANAGER
  );

  const oldLast = sheet.getLastRow();


  const cycles = getCycleRows_().filter(function (row) {

    return String(
      row[CFG.CYCLE.STATUS - 1] || ''
    ).toLowerCase() === 'pending';

  });


  const now = new Date();


  cycles.sort(function (a, b) {

    const da = combineDateAndTime_(
      a[CFG.CYCLE.DUE - 1],
      a[CFG.CYCLE.DUE_TIME - 1]
    );

    const db = combineDateAndTime_(
      b[CFG.CYCLE.DUE - 1],
      b[CFG.CYCLE.DUE_TIME - 1]
    );

    return (
      (da ? da.getTime() : 0) -
      (db ? db.getTime() : 0)
    );

  });


  const output = cycles.map(function (row) {

    const dueAt = combineDateAndTime_(
      row[CFG.CYCLE.DUE - 1],
      row[CFG.CYCLE.DUE_TIME - 1]
    );

    return [
      row[CFG.CYCLE.ID - 1],
      row[CFG.CYCLE.OWNER - 1],
      row[CFG.CYCLE.TASK - 1],
      row[CFG.CYCLE.DUE - 1],
      row[CFG.CYCLE.DUE_TIME - 1],
      row[CFG.CYCLE.STATUS - 1],
      dueState_(dueAt, now),
      row[CFG.CYCLE.MANAGER_ALERT_AT - 1],
      isDate_(row[CFG.CYCLE.MANAGER_ACK_AT - 1])
        ? 'Yes'
        : 'No',
      row[CFG.CYCLE.MANAGER_ACK_BY - 1],
      row[CFG.CYCLE.MANAGER_ACK_AT - 1],
      row[CFG.CYCLE.LEADERSHIP_AT - 1],
      row[CFG.CYCLE.REQUIRED_PROOF - 1],
      row[CFG.CYCLE.PROOF - 1]
    ];

  });


  if (oldLast >= CFG.MANAGER_DATA_ROW) {

    sheet.getRange(
      CFG.MANAGER_DATA_ROW,
      1,
      oldLast - CFG.MANAGER_DATA_ROW + 1,
      14
    ).clearContent();

  }


  if (output.length) {

    sheet.getRange(
      CFG.MANAGER_DATA_ROW,
      1,
      output.length,
      14
    ).setValues(output);

    sheet.getRange(
      CFG.MANAGER_DATA_ROW,
      4,
      output.length,
      1
    ).setNumberFormat('dd mmm yyyy');

    sheet.getRange(
      CFG.MANAGER_DATA_ROW,
      8,
      output.length,
      1
    ).setNumberFormat('dd mmm yyyy hh:mm');

    sheet.getRange(
      CFG.MANAGER_DATA_ROW,
      11,
      output.length,
      2
    ).setNumberFormat('dd mmm yyyy hh:mm');

  }

}


/**
 * Reads Task Master and returns active and inactive tasks
 * keyed by Task ID.
 */
function getTaskMap_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.TASKS);

  const last = sheet.getLastRow();

  const map = {};

  if (last < CFG.DATA_ROW) return map;


  const rows = sheet.getRange(
    CFG.DATA_ROW,
    1,
    last - CFG.DATA_ROW + 1,
    12
  ).getValues();


  rows.forEach(function (row) {

    const id = String(
      row[0] || ''
    ).trim();

    if (!id) return;

    map[id] = {

      taskId: id,

      taskName: String(row[1] || ''),

      owner: String(row[2] || ''),

      frequency: String(row[3] || ''),

      intervalDays: numberOr_(row[4], 1),

      weekdaysOnly:
        String(row[5] || '').toLowerCase() === 'yes',

      dueTime: row[6] || '17:00',

      taskType:
        String(row[7] || 'Other').toLowerCase(),

      requiredProof: String(row[8] || ''),

      active:
        String(row[9] || '').toLowerCase() === 'yes',

      firstDueDate:
        isDate_(row[10])
          ? dateOnly_(row[10])
          : null,

      notes: String(row[11] || '')

    };

  });

  return map;

}


function getTaskType_(taskId) {

  const task = getTaskMap_()[taskId];

  return task
    ? task.taskType
    : 'other';

}


/**
 * Returns all populated rows from the Cycle Tracker.
 * Always reads exactly 21 columns (A:U).
 */
function getCycleRows_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.CYCLES);

  const last = sheet.getLastRow();

  if (last < CFG.DATA_ROW) return [];

  return sheet.getRange(
    CFG.DATA_ROW,
    1,
    last - CFG.DATA_ROW + 1,
    21
  ).getValues().filter(function (row) {

    return String(row[0] || '').trim() !== '';

  });

}


function hasPendingCycleForTask_(taskId) {

  return getCycleRows_().some(function (row) {

    return (
      String(row[CFG.CYCLE.TASK_ID - 1]) === taskId &&
      String(row[CFG.CYCLE.STATUS - 1]).toLowerCase() === 'pending'
    );

  });

}


function getNextCycleNumber_(taskId) {

  let maxNumber = 0;

  getCycleRows_().forEach(function (row) {

    if (
      String(row[CFG.CYCLE.TASK_ID - 1]) === taskId
    ) {

      maxNumber = Math.max(
        maxNumber,
        numberOr_(
          row[CFG.CYCLE.NUMBER - 1],
          0
        )
      );

    }

  });

  return maxNumber + 1;

}


/**
 * Adds a new cycle without changing previous cycle history.
 * Writes exactly 21 columns.
 */
function appendCycle_(
  task,
  cycleNumber,
  startDate,
  dueDate
) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.CYCLES);

  const start = dateOnly_(
    startDate || new Date()
  );

  const due = dateOnly_(
    dueDate || start
  );

  const cycleId = makeCycleId_(
    task.taskId,
    due,
    cycleNumber
  );

  if (findCycleRow_(cycleId)) return;


  const newRow = [

    cycleId,                         // A: Cycle ID
    task.taskId,                     // B: Task ID
    task.taskName,                   // C: Task Name
    task.owner,                      // D: Assigned To
    cycleNumber,                     // E: Cycle #
    start,                           // F: Cycle Start Date
    due,                             // G: Due Date
    task.dueTime || '17:00',         // H: Due Time
    'Pending',                       // I: Status
    task.requiredProof || '',        // J: Required Proof
    '',                              // K: Ticket ID / Completion Note
    '',                              // L: Proof Link
    '',                              // M: Completed By
    '',                              // N: Completed At
    '',                              // O: Manager Alert Sent At
    '',                              // P: Manager Reviewed By
    '',                              // Q: Manager Reviewed At
    '',                              // R: Leadership Escalated At
    new Date(),                      // S: Last Updated At
    task.notes || '',                // T: Notes
    ''                               // U: Identity Source

  ];


  sheet.appendRow(newRow);

  const newSheetRow = sheet.getLastRow();

  // Copy the existing row formatting into the new cycle.

  sheet.getRange(
    CFG.DATA_ROW,
    1,
    1,
    21
  ).copyTo(
    sheet.getRange(
      newSheetRow,
      1,
      1,
      21
    ),
    { formatOnly: true }
  );

}


function makeCycleId_(
  taskId,
  dueDate,
  cycleNumber
) {

  const safeTask = String(taskId)
    .replace(/[^A-Za-z0-9]+/g, '')
    .toUpperCase();

  const dateTag = Utilities.formatDate(
    dateOnly_(dueDate),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  return (
    'CYC-' +
    safeTask +
    '-' +
    dateTag +
    '-' +
    String(cycleNumber).padStart(3, '0')
  );

}


/**
 * Daily tasks move to the next working day.
 * Other recurring tasks restart from completion date + interval.
 */
function calculateNextDueDate_(
  completedAt,
  task
) {

  const next = dateOnly_(completedAt);

  if (task.weekdaysOnly) {

    next.setDate(next.getDate() + 1);

    while (
      next.getDay() === 0 ||
      next.getDay() === 6
    ) {

      next.setDate(next.getDate() + 1);

    }

    return next;

  }

  next.setDate(
    next.getDate() +
    Math.max(1, task.intervalDays)
  );

  return next;

}


/* =========================================================
   6. COMPLETION HISTORY AND CYCLE LOOKUPS
   ========================================================= */


/**
 * Adds one completion record per cycle.
 */
function appendCompletionLog_(
  cycle,
  completedBy,
  completedAt,
  proof,
  proofLink,
  nextDue,
  identitySource
) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.COMPLETIONS);

  const cycleId = String(
    cycle[CFG.CYCLE.ID - 1]
  );

  if (completionAlreadyLogged_(cycleId)) return;


  sheet.appendRow([

    cycleId,

    cycle[CFG.CYCLE.TASK_ID - 1],

    cycle[CFG.CYCLE.TASK - 1],

    cycle[CFG.CYCLE.OWNER - 1],

    completedBy,

    completedAt,

    proof,

    proofLink,

    cycle[CFG.CYCLE.NUMBER - 1],

    nextDue || '',

    identitySource

  ]);

}


/**
 * Checks whether a completion record already exists.
 */
function completionAlreadyLogged_(cycleId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.COMPLETIONS);

  const last = sheet.getLastRow();

  if (last < CFG.DATA_ROW) return false;

  const ids = sheet.getRange(
    CFG.DATA_ROW,
    1,
    last - CFG.DATA_ROW + 1,
    1
  ).getValues();

  return ids.some(function (row) {

    return String(
      row[0] || ''
    ).trim() === cycleId;

  });

}


/**
 * Finds the Cycle Tracker row for a given Cycle ID.
 */
function findCycleRow_(cycleId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.CYCLES);

  const last = sheet.getLastRow();

  if (last < CFG.DATA_ROW) return 0;

  const ids = sheet.getRange(
    CFG.DATA_ROW,
    CFG.CYCLE.ID,
    last - CFG.DATA_ROW + 1,
    1
  ).getValues();

  for (let i = 0; i < ids.length; i++) {

    if (
      String(ids[i][0] || '').trim() === cycleId
    ) {

      return CFG.DATA_ROW + i;

    }

  }

  return 0;

}


/* =========================================================
   7. EMAIL AND SETTINGS
   ========================================================= */


/**
 * Gets a real email address for a team member.
 * Addresses ending in example.com are ignored.
 */
function getTeamEmail_(name) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.TEAM);

  const last = sheet.getLastRow();

  if (last < CFG.DATA_ROW) return '';

  const rows = sheet.getRange(
    CFG.DATA_ROW,
    1,
    last - CFG.DATA_ROW + 1,
    4
  ).getValues();

  const found = rows.find(function (row) {

    return (
      String(row[0] || '').trim().toLowerCase() ===
      String(name || '').trim().toLowerCase()
    );

  });

  const email = found
    ? String(found[2] || '').trim()
    : '';

  return isUsableEmail_(email)
    ? email
    : '';

}


/**
 * Gets the active Google account email when Google exposes it.
 */
function getActiveEmail_() {

  try {

    return String(
      Session.getActiveUser().getEmail() || ''
    ).trim();

  } catch (err) {

    return '';

  }

}


/**
 * Reads the SLA & Settings tab.
 */
function getSettings_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.SETTINGS);

  const last = sheet.getLastRow();

  const rowCount = Math.max(
    0,
    last - CFG.DATA_ROW + 1
  );

  const rows = rowCount
    ? sheet.getRange(
        CFG.DATA_ROW,
        1,
        rowCount,
        3
      ).getValues()
    : [];

  const map = {};


  rows.forEach(function (row) {

    const key = String(
      row[0] || ''
    ).trim();

    if (key) {
      map[key] = row[1];
    }

  });


  const leadershipText = String(
    map['Leadership email(s)'] || ''
  );

  const managerCandidate = String(
    map['Manager email'] || ''
  ).trim();


  return {

    managerEmail: isUsableEmail_(
      managerCandidate
    ) ? managerCandidate : '',

    leadershipEmails: leadershipText
      .split(/[;,]/)
      .map(function (email) {
        return email.trim();
      })
      .filter(isUsableEmail_),

    dryRun:
      String(map['Dry Run Mode'] || 'TRUE')
        .trim()
        .toUpperCase() === 'TRUE',

    managerAckHours:
      map['Manager acknowledgement SLA (business hours)'],

    resolutionHours:
      map['Task resolution SLA (calendar hours)'],

    businessStart:
      String(map['Business day starts'] || '09:00'),

    businessEnd:
      String(map['Business day ends'] || '18:00')

  };

}


/**
 * Creates and sends the manager or leadership email.
 * With Dry Run Mode TRUE, it only records the simulated alert.
 */
function sendAlert_(
  type,
  cycle,
  settings,
  now
) {

  const owner =
    cycle[CFG.CYCLE.OWNER - 1];

  const task =
    cycle[CFG.CYCLE.TASK - 1];

  const dueDate =
    cycle[CFG.CYCLE.DUE - 1];

  const dueTime =
    cycle[CFG.CYCLE.DUE_TIME - 1];

  const ticket =
    cycle[CFG.CYCLE.PROOF - 1];


  const isManager = type === 'Manager';

  const recipients = isManager
    ? [settings.managerEmail].filter(Boolean)
    : settings.leadershipEmails;


  const subject = isManager
    ? '[Task Hub] Task overdue: ' + task
    : '[Task Hub] Escalation needed: ' + task;


  const body = [

    'Hello,',

    '',

    isManager
      ? 'This task has passed its due time and is still open.'
      : 'This task is still open and has crossed the follow-up SLA.',

    '',

    'Task: ' + task,

    'Assigned to: ' + owner,

    'Due: ' + formatDateTime_(dueDate, dueTime),

    'Cycle ID: ' + cycle[CFG.CYCLE.ID - 1],

    'Required proof: ' +
      (cycle[CFG.CYCLE.REQUIRED_PROOF - 1] ||
       'Completion update'),

    'Latest completion note / ticket: ' +
      (ticket || 'Not submitted yet'),

    '',

    isManager
      ? 'Please review the issue and acknowledge it in the Manager Dashboard. The employee still needs to submit the completion update.'
      : 'Please help unblock the issue. The manager acknowledgement or resolution window has been missed.',

    '',

    'This note was generated by the Task Hub recurring task automation.'

  ].join('\n');


  // Dry Run Mode: record the alert without sending an email.

  if (settings.dryRun) {

    return {
      stamp: true,
      outcome: 'DRY RUN — no email sent',
      summary: subject + ' (dry run)'
    };

  }


  const validRecipients = recipients.filter(
    isUsableEmail_
  );


  if (!validRecipients.length) {

    return {
      stamp: true,
      outcome: 'NOT SENT — update recipient email',
      summary: subject +
        ' (recipient missing or still a placeholder)'
    };

  }


  try {

    MailApp.sendEmail({

      to: validRecipients.join(','),

      subject: subject,

      body: body,

      name: 'Task Hub Automation'

    });

    return {
      stamp: true,
      outcome: 'Sent',
      summary: subject
    };

  } catch (err) {

    return {
      stamp: false,
      outcome: 'SEND ERROR — ' + err.message,
      summary: subject
    };

  }

}


/**
 * Rejects placeholder email addresses such as example.com.
 */
function isUsableEmail_(email) {

  const value = String(
    email || ''
  ).trim();

  return (
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value) &&
    !/example\.com$/i.test(value)
  );

}


/**
 * Appends an alert event to Alert Log.
 */
function logAlert_(
  cycleId,
  type,
  recipient,
  at,
  outcome,
  summary
) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CFG.SHEETS.ALERTS);

  sheet.appendRow([

    Utilities.getUuid(),

    cycleId,

    type,

    recipient || '(not configured)',

    at,

    outcome,

    summary

  ]);

}


/* =========================================================
   8. SLA TIME CALCULATIONS
   ========================================================= */


/**
 * Adds business hours to a timestamp.
 * Saturday and Sunday are excluded.
 * Only the configured business hours are counted.
 */
function addBusinessHours_(
  startDate,
  hours,
  settings
) {

  const cursor = new Date(startDate);

  const minutesNeeded = Math.max(
    0,
    Math.round(hours * 60)
  );

  let counted = 0;

  const startParts = parseClock_(
    settings.businessStart || '09:00'
  );

  const endParts = parseClock_(
    settings.businessEnd || '18:00'
  );


  while (counted < minutesNeeded) {

    cursor.setMinutes(
      cursor.getMinutes() + 1
    );

    const day = cursor.getDay();

    const minuteOfDay =
      cursor.getHours() * 60 +
      cursor.getMinutes();

    const withinBusinessHours =
      day !== 0 &&
      day !== 6 &&
      minuteOfDay >=
        startParts.hour * 60 + startParts.minute &&
      minuteOfDay <
        endParts.hour * 60 + endParts.minute;

    if (withinBusinessHours) {
      counted++;
    }

  }

  return cursor;

}


/**
 * Converts a time value to hours and minutes.
 */
function parseClock_(value) {

  if (value instanceof Date) {

    return {
      hour: value.getHours(),
      minute: value.getMinutes()
    };

  }

  const parts = String(
    value || '17:00'
  ).split(':');

  return {

    hour: Number(parts[0] || 0),

    minute: Number(parts[1] || 0)

  };

}


/**
 * Combines a spreadsheet date and time into one Date object.
 */
function combineDateAndTime_(
  dateValue,
  timeValue
) {

  if (!isDate_(dateValue)) return null;

  const result = new Date(dateValue);

  result.setHours(0, 0, 0, 0);

  const clock = parseClock_(timeValue);

  result.setHours(
    clock.hour,
    clock.minute,
    0,
    0
  );

  return result;

}


/**
 * Labels a task as overdue, due today or upcoming.
 */
function dueState_(
  dueAt,
  now
) {

  if (!dueAt) {
    return 'Check due date';
  }

  if (
    now.getTime() >= dueAt.getTime()
  ) {
    return 'Overdue';
  }

  const today = dateOnly_(now).getTime();

  return (
    dateOnly_(dueAt).getTime() === today
  )
    ? 'Due today'
    : 'Upcoming';

}


/**
 * Removes the time portion from a date.
 */
function dateOnly_(value) {

  const d = new Date(value);

  return new Date(
    d.getFullYear(),
    d.getMonth(),
    d.getDate()
  );

}


/**
 * Checks whether a value is a valid JavaScript Date.
 */
function isDate_(value) {

  return (
    value instanceof Date &&
    !isNaN(value.getTime())
  );

}


/**
 * Converts a value to a number or uses a fallback.
 */
function numberOr_(
  value,
  fallback
) {

  const n = Number(value);

  return isFinite(n)
    ? n
    : fallback;

}


/**
 * Formats a date and time for email notifications.
 */
function formatDateTime_(
  dateValue,
  timeValue
) {

  const dateText = isDate_(dateValue)

    ? Utilities.formatDate(
        dateValue,
        Session.getScriptTimeZone(),
        'dd MMM yyyy'
      )

    : String(dateValue || '');

  return (
    dateText +
    ' at ' +
    String(timeValue || '17:00')
  );

}
