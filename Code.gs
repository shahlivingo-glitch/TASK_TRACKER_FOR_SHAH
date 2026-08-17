var SHEET_NAMES = {
  EMPLOYEES: 'Employees',
  TASKS: 'Tasks',
  WORKLOG: 'WorkLog',
  WORKSUMMARY: 'WorkSummary'
};

var SHEET_HEADERS = {
  Employees: ['EmpID', 'Name', 'PIN', 'Active'],
  Tasks: ['TaskID', 'EmpID', 'Category', 'TaskName', 'TargetQty', 'AssignedBy', 'AssignedDate', 'Status'],
  WorkLog: ['LogID', 'TaskID', 'EmpID', 'Action', 'Timestamp'],
  WorkSummary: ['TaskID', 'EmpID', 'Name', 'Category', 'TaskName', 'TargetQty', 'QtyDone', 'TotalWorkMins', 'TotalBreakMins', 'StartedAt', 'FinishedAt', 'Status', 'Note']
};

var OLD_TASKS_HEADER = ['TaskID', 'EmpID', 'WorkType', 'TargetQty', 'AssignedBy', 'AssignedDate', 'Status'];
var OLD_WORKSUMMARY_HEADER = ['TaskID', 'EmpID', 'Name', 'WorkType', 'TargetQty', 'QtyDone', 'TotalWorkMins', 'TotalBreakMins', 'StartedAt', 'FinishedAt', 'Status', 'Note'];

function doGet(e) {
  setupSheets();
  var params = (e && e.parameter) || {};
  var action = params.action;

  var handlers = {
    getEmployees: function () { return getEmployees(); },
    getMyTasks: function () { return getMyTasks(params.empId); },
    getAdminDashboard: function () { return getAdminDashboard(); },
    authenticateEmployee: function () { return authenticateEmployee(params.pin); },
    authenticateAdmin: function () { return authenticateAdmin(params.pin); },
    startWork: function () { return startWork(params.taskId, params.empId); },
    takeBreak: function () { return takeBreak(params.taskId, params.empId); },
    resumeWork: function () { return resumeWork(params.taskId, params.empId); },
    stopWork: function () { return stopWork(params.taskId, params.empId, Number(params.qtyDone), params.note || ''); },
    assignTasks: function () { return assignTasks(params.empId, params.category, JSON.parse(params.tasks), params.assignedBy); },
    deleteTask: function () { return deleteTask(params.taskId); }
  };

  var ok = true;
  var data;
  try {
    var handler = handlers[action];
    if (!handler) throw new Error('Unknown action: ' + action);
    data = handler();
  } catch (err) {
    ok = false;
    data = { message: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Work Tracker').addItem('Initialize Sheets', 'setupSheets').addToUi();
}

function setupSheets() {
  withLock_(function () {
    Object.keys(SHEET_HEADERS).forEach(function (name) {
      getSheet_(name);
    });
    migrateHeader_(SHEET_NAMES.TASKS, OLD_TASKS_HEADER, SHEET_HEADERS.Tasks, 2, 'General');
    migrateHeader_(SHEET_NAMES.WORKSUMMARY, OLD_WORKSUMMARY_HEADER, SHEET_HEADERS.WorkSummary, 3, 'General');
  });
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS[name]);
  }
  return sheet;
}

function migrateHeader_(sheetName, oldHeader, newHeader, insertAfterCol, fillValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) return;
  var currentHeader = sheet.getRange(1, 1, 1, oldHeader.length).getValues()[0];
  var isOld = oldHeader.every(function (h, i) { return currentHeader[i] === h; });
  if (!isOld) return;

  sheet.insertColumnAfter(insertAfterCol);
  var dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    sheet.getRange(2, insertAfterCol + 1, dataRows, 1).setValue(fillValue);
  }
  sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      obj[h] = row[i];
    });
    return obj;
  });
}

function generateId_(prefix) {
  return prefix + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

function findRowIndex_(sheet, colName, value) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var colIndex = headers.indexOf(colName);
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][colIndex]) === String(value)) {
      return i + 1;
    }
  }
  return -1;
}

function toISO_(d) {
  return d instanceof Date ? d.toISOString() : d;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function logWork_(taskId, empId, action) {
  var sheet = getSheet_(SHEET_NAMES.WORKLOG);
  var logId = generateId_('L');
  sheet.appendRow([logId, taskId, empId, action, new Date()]);
}

function setTaskStatus_(taskId, status) {
  var sheet = getSheet_(SHEET_NAMES.TASKS);
  var rowIndex = findRowIndex_(sheet, 'TaskID', taskId);
  if (rowIndex === -1) {
    throw new Error('Task not found: ' + taskId);
  }
  var statusCol = SHEET_HEADERS.Tasks.indexOf('Status') + 1;
  sheet.getRange(rowIndex, statusCol).setValue(status);
}

function getEmployees() {
  var employees = sheetToObjects_(getSheet_(SHEET_NAMES.EMPLOYEES));
  return employees
    .filter(function (e) {
      return String(e.Active).toUpperCase() === 'Y';
    })
    .map(function (e) {
      return { empId: e.EmpID, name: e.Name };
    });
}

function assignTasks(empId, category, tasks, assignedBy) {
  return withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.TASKS);
    var now = new Date();
    var taskIds = tasks.map(function (t, i) {
      var taskId = generateId_('T') + '_' + i;
      sheet.appendRow([taskId, empId, category, t.taskName, Number(t.targetQty), assignedBy, now, 'Pending']);
      return taskId;
    });
    return { taskIds: taskIds };
  });
}

function deleteTask(taskId) {
  return withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.TASKS);
    var rowIndex = findRowIndex_(sheet, 'TaskID', taskId);
    if (rowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }
    sheet.deleteRow(rowIndex);
    return { success: true };
  });
}

function getMyTasks(empId) {
  var tasks = sheetToObjects_(getSheet_(SHEET_NAMES.TASKS));
  return tasks
    .filter(function (t) {
      return String(t.EmpID) === String(empId) && t.Status !== 'Completed';
    })
    .map(function (t) {
      return {
        taskId: t.TaskID,
        category: t.Category,
        taskName: t.TaskName,
        targetQty: t.TargetQty,
        status: t.Status,
        assignedDate: toISO_(t.AssignedDate)
      };
    });
}

function startWork(taskId, empId) {
  return withLock_(function () {
    logWork_(taskId, empId, 'Start');
    setTaskStatus_(taskId, 'InProgress');
    return { success: true };
  });
}

function takeBreak(taskId, empId) {
  return withLock_(function () {
    logWork_(taskId, empId, 'Break');
    setTaskStatus_(taskId, 'OnBreak');
    return { success: true };
  });
}

function resumeWork(taskId, empId) {
  return withLock_(function () {
    logWork_(taskId, empId, 'Resume');
    setTaskStatus_(taskId, 'InProgress');
    return { success: true };
  });
}

function computeWorkSummary_(logs) {
  logs.sort(function (a, b) {
    return new Date(a.Timestamp) - new Date(b.Timestamp);
  });

  var totalWorkMs = 0;
  var totalBreakMs = 0;
  var openWorkStart = null;
  var openBreakStart = null;
  var startedAt = null;
  var finishedAt = null;

  logs.forEach(function (log) {
    var t = new Date(log.Timestamp);
    if (log.Action === 'Start') {
      if (!startedAt) startedAt = t;
      openWorkStart = t;
    } else if (log.Action === 'Resume') {
      if (openBreakStart) {
        totalBreakMs += t - openBreakStart;
        openBreakStart = null;
      }
      openWorkStart = t;
    } else if (log.Action === 'Break') {
      if (openWorkStart) {
        totalWorkMs += t - openWorkStart;
        openWorkStart = null;
      }
      openBreakStart = t;
    } else if (log.Action === 'Stop') {
      if (openWorkStart) {
        totalWorkMs += t - openWorkStart;
        openWorkStart = null;
      }
      if (openBreakStart) {
        totalBreakMs += t - openBreakStart;
        openBreakStart = null;
      }
      finishedAt = t;
    }
  });

  return {
    totalWorkMins: Math.round(totalWorkMs / 60000),
    totalBreakMins: Math.round(totalBreakMs / 60000),
    startedAt: startedAt,
    finishedAt: finishedAt
  };
}

function stopWork(taskId, empId, qtyDone, note) {
  return withLock_(function () {
    logWork_(taskId, empId, 'Stop');

    var taskSheet = getSheet_(SHEET_NAMES.TASKS);
    var taskRowIndex = findRowIndex_(taskSheet, 'TaskID', taskId);
    if (taskRowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }
    var taskRowValues = taskSheet.getRange(taskRowIndex, 1, 1, SHEET_HEADERS.Tasks.length).getValues()[0];
    var task = {};
    SHEET_HEADERS.Tasks.forEach(function (h, i) {
      task[h] = taskRowValues[i];
    });

    var employees = sheetToObjects_(getSheet_(SHEET_NAMES.EMPLOYEES));
    var emp = employees.filter(function (e) {
      return String(e.EmpID) === String(empId);
    })[0];
    var empName = emp ? emp.Name : '';

    var allLogs = sheetToObjects_(getSheet_(SHEET_NAMES.WORKLOG)).filter(function (l) {
      return String(l.TaskID) === String(taskId);
    });
    allLogs.sort(function (a, b) {
      return new Date(a.Timestamp) - new Date(b.Timestamp);
    });
    var prevStopIdx = -1;
    for (var i = 0; i < allLogs.length - 1; i++) {
      if (allLogs[i].Action === 'Stop') prevStopIdx = i;
    }
    var sessionLogs = prevStopIdx === -1 ? allLogs : allLogs.slice(prevStopIdx + 1);
    var summary = computeWorkSummary_(sessionLogs);

    var qty = Number(qtyDone) || 0;
    var target = Number(task.TargetQty) || 0;
    var isComplete = qty >= target;
    var finalStatus = isComplete ? 'Completed' : 'Partial';

    getSheet_(SHEET_NAMES.WORKSUMMARY).appendRow([
      taskId,
      empId,
      empName,
      task.Category,
      task.TaskName,
      target,
      qty,
      summary.totalWorkMins,
      summary.totalBreakMins,
      summary.startedAt,
      summary.finishedAt,
      finalStatus,
      note || ''
    ]);

    if (isComplete) {
      taskSheet.deleteRow(taskRowIndex);
    } else {
      var remaining = target - qty;
      var qtyCol = SHEET_HEADERS.Tasks.indexOf('TargetQty') + 1;
      var statusCol = SHEET_HEADERS.Tasks.indexOf('Status') + 1;
      taskSheet.getRange(taskRowIndex, qtyCol).setValue(remaining);
      taskSheet.getRange(taskRowIndex, statusCol).setValue('Pending');
    }

    return { success: true, status: finalStatus };
  });
}

function getAdminDashboard() {
  var tasks = sheetToObjects_(getSheet_(SHEET_NAMES.TASKS));
  var employees = sheetToObjects_(getSheet_(SHEET_NAMES.EMPLOYEES));
  var empMap = {};
  employees.forEach(function (e) {
    empMap[e.EmpID] = e.Name;
  });

  return tasks
    .map(function (t) {
      return {
        taskId: t.TaskID,
        empId: t.EmpID,
        empName: empMap[t.EmpID] || t.EmpID,
        category: t.Category,
        taskName: t.TaskName,
        targetQty: t.TargetQty,
        status: t.Status,
        assignedDate: toISO_(t.AssignedDate)
      };
    })
    .sort(function (a, b) {
      return new Date(b.assignedDate) - new Date(a.assignedDate);
    });
}

function authenticateAdmin(pin) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!stored) {
    throw new Error('Admin PIN not configured. In the Apps Script editor, go to Project Settings > Script Properties and add a property named ADMIN_PIN.');
  }
  return String(pin) === String(stored);
}

function authenticateEmployee(pin) {
  var employees = sheetToObjects_(getSheet_(SHEET_NAMES.EMPLOYEES));
  var match = employees.filter(function (e) {
    return String(e.PIN) === String(pin) && String(e.Active).toUpperCase() === 'Y';
  })[0];
  if (!match) return null;
  return { empId: match.EmpID, name: match.Name };
}
