// ============================================================
// ENTRY POINTS
// ============================================================

function getJobsDebug() {
  var sheet = getSheet('Jobs');
  var lastRow = sheet ? sheet.getLastRow() : -1;
  var jobs = getJobs();
  return { sheetExists: !!sheet, lastRow: lastRow, jobsLength: jobs.length, first: jobs[0] || null };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('HaiLux - Quản Lý Công Việc')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// SETUP
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = {
    Users:         ['id','name','role','email','created_at'],
    Jobs:          ['id','code','name','category','customer_name','customer_contact','received_date','deadline','revenue','repair_scope','status_id','notes','created_at','avatar_id'],
    Tasks:         ['id','job_id','name','order','assignee_id','deadline','status_id','completed_at','notes','created_at','evidence_id','emp_notes'],
    Statuses:      ['id','entity_type','label','color','order'],
    TaskTemplates: ['id','name','description'],
    Settings:      ['key','value'],
  };

  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      styleHeaderRow(sheet);
    }
  }

  seedDefaultData(ss);
  return 'Setup hoàn tất!';
}

function styleHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground('#0d1b4b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function seedDefaultData(ss) {
  const statusSheet = ss.getSheetByName('Statuses');
  if (statusSheet.getLastRow() <= 1) {
    const now = new Date().toISOString();
    const statuses = [
      [genId(), 'job',  'Chưa làm',   '#9E9E9E', 1],
      [genId(), 'job',  'Đang làm',   '#2196F3', 2],
      [genId(), 'job',  'Hoàn thành', '#4CAF50', 3],
      [genId(), 'job',  'Hủy',        '#F44336', 4],
      [genId(), 'task', 'Chưa làm',   '#9E9E9E', 1],
      [genId(), 'task', 'Đang làm',   '#FF9800', 2],
      [genId(), 'task', 'Hoàn thành', '#4CAF50', 3],
      [genId(), 'task', 'Treo',       '#9C27B0', 4],
    ];
    statuses.forEach(r => statusSheet.appendRow(r));
  }

  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['manager_password', '1234']);
    settingsSheet.appendRow(['wholesale_password', '0000']);
    settingsSheet.appendRow(['app_name', 'HaiLux']);
    settingsSheet.appendRow(['drive_folder_id', '']);
  }

  const tmplSheet = ss.getSheetByName('TaskTemplates');
  if (tmplSheet.getLastRow() <= 1) {
    const templates = ['Tách túi','Dựng form','Đi viền','Thay da','Lắp ráp vào','Vệ sinh','Mạ khóa','Dán seal','Phục hồi màu','Xử lý xước'];
    templates.forEach((name, i) => tmplSheet.appendRow([genId(), name, '']));
  }
}

// ============================================================
// UTILITIES
// ============================================================

function genId() {
  return Utilities.getUuid();
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script chưa được gắn vào Google Sheet.');
  return ss.getSheetByName(name);
}

// Dùng cho write operations — throw nếu sheet không tồn tại
function requireSheet(name) {
  const sheet = getSheet(name);
  if (!sheet) throw new Error('Sheet "' + name + '" không tồn tại. Vui lòng chạy setupSheets() trước.');
  return sheet;
}

function sheetToObjects(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      var val = row[i];
      if (val instanceof Date) val = val.toISOString();
      obj[h] = val;
    });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ============================================================
// AUTH
// ============================================================

function verifyWholesalePassword(password) {
  const settings = sheetToObjects(getSheet('Settings'));
  const entry = settings.find(s => s.key === 'wholesale_password');
  return entry && String(entry.value) === String(password);
}

function getWholesaleJobs() {
  return getJobs().filter(j => (j.category||'').toLowerCase().includes('sỉ'));
}

function verifyManagerPassword(password) {
  const settings = sheetToObjects(getSheet('Settings'));
  const entry = settings.find(s => s.key === 'manager_password');
  return entry && String(entry.value) === String(password);
}

// Thêm cột mới vào sheet đã có, không xóa dữ liệu
function migrateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Thêm cột vào sheet đã có nếu thiếu
  const migrations = {
    'Jobs':  ['avatar_id'],
    'Tasks': ['evidence_id', 'emp_notes'],
  };

  Object.entries(migrations).forEach(([name, cols]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    cols.forEach(col => {
      if (!headers.includes(col)) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(col);
        sheet.getRange(1, nextCol).setBackground('#0d1b4b').setFontColor('#fff').setFontWeight('bold');
        Logger.log('Added column: ' + col + ' to ' + name);
      } else {
        Logger.log('Column already exists: ' + col + ' in ' + name);
      }
    });
  });

  // Thêm setting wholesale_password nếu chưa có
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet) {
    const data = settingsSheet.getDataRange().getValues();
    const keys = data.map(r => r[0]);
    if (!keys.includes('wholesale_password')) {
      settingsSheet.appendRow(['wholesale_password', '0000']);
      Logger.log('Added wholesale_password setting');
    } else {
      Logger.log('wholesale_password already exists');
    }
  }

  return 'Migration hoàn tất! Xem Nhật ký thực thi để biết chi tiết.';
}

function cleanEmptyTasks() {
  const sheet = getSheet('Tasks');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('name');
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][nameCol] || !String(data[i][nameCol]).trim()) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return 'Đã xóa ' + deleted + ' task rỗng';
}

function debugAll() {
  var results = {};
  var sheets = ['Users','Jobs','Tasks','Statuses','TaskTemplates','Settings'];
  sheets.forEach(function(name) {
    try {
      var sheet = getSheet(name);
      if (!sheet) { results[name] = 'MISSING'; return; }
      results[name] = sheet.getLastRow() + ' rows';
    } catch(e) { results[name] = 'ERROR: ' + e.message; }
  });
  try { results['_getInitialData'] = JSON.stringify(getInitialData()).substring(0, 200); } catch(e) { results['_getInitialData'] = 'ERROR: ' + e.message; }
  try { results['_getDashboardStats'] = 'OK: ' + JSON.stringify(getDashboardStats()).substring(0,100); } catch(e) { results['_getDashboardStats'] = 'ERROR: ' + e.message; }
  return JSON.stringify(results, null, 2);
}

// ============================================================
// USERS API
// ============================================================

function getUsers() {
  return sheetToObjects(getSheet('Users'));
}

function createUser(data) {
  const sheet = getSheet('Users');
  const id = genId();
  const now = new Date().toISOString();
  sheet.appendRow([id, data.name, data.role || 'employee', data.email || '', now]);
  return { id, ...data, created_at: now };
}

function updateUser(id, data) {
  const sheet = getSheet('Users');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteUser(id) {
  const sheet = getSheet('Users');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

// ============================================================
// STATUSES API
// ============================================================

function getStatuses() {
  return sheetToObjects(getSheet('Statuses'));
}

function createStatus(data) {
  const sheet = getSheet('Statuses');
  const id = genId();
  sheet.appendRow([id, data.entity_type, data.label, data.color, data.order]);
  return { id, ...data };
}

function updateStatus(id, data) {
  const sheet = getSheet('Statuses');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteStatus(id) {
  const sheet = getSheet('Statuses');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

// ============================================================
// TASK TEMPLATES API
// ============================================================

function getTaskTemplates() {
  return sheetToObjects(getSheet('TaskTemplates'));
}

function createTaskTemplate(data) {
  const sheet = getSheet('TaskTemplates');
  const id = genId();
  sheet.appendRow([id, data.name, data.description || '']);
  return { id, ...data };
}

function deleteTaskTemplate(id) {
  const sheet = getSheet('TaskTemplates');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

// ============================================================
// SETTINGS API
// ============================================================

function getSettings() {
  return sheetToObjects(getSheet('Settings'));
}

function updateSetting(key, value) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  sheet.appendRow([key, value]);
  return true;
}

// ============================================================
// JOBS API
// ============================================================

function getJobs() {
  return sheetToObjects(getSheet('Jobs'));
}

function getJob(id) {
  return getJobs().find(j => j.id === id) || null;
}

function getJobsWithStats() {
  const jobs  = getJobs();
  const tasks = getTasks();
  return jobs.map(j => {
    const jobTasks  = tasks.filter(t => t.job_id === j.id);
    const assignees = [...new Set(jobTasks.map(t => t.assignee_id).filter(Boolean))];
    return { ...j, task_count: jobTasks.length, assignee_count: assignees.length };
  });
}

function createJob(data) {
  const sheet = getSheet('Jobs');
  const id = genId();
  const now = new Date().toISOString();
  const statusId = data.status_id || getDefaultStatusId('job');
  sheet.appendRow([
    id,
    data.code || '',
    data.name || '',
    data.category || '',
    data.customer_name || '',
    data.customer_contact || '',
    data.received_date || '',
    data.deadline || '',
    data.revenue || 0,
    data.repair_scope || '',
    statusId,
    data.notes || '',
    now,
  ]);
  return { id, ...data, status_id: statusId, created_at: now };
}

function updateJob(id, data) {
  const sheet = getSheet('Jobs');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);
  Object.keys(data).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });
  return { id, ...data };
}

function deleteJob(id) {
  const sheet = getSheet('Jobs');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  // also delete tasks and images
  const tasks = getTasks().filter(t => t.job_id === id);
  tasks.forEach(t => deleteTask(t.id));
  return true;
}

function getDefaultStatusId(entityType) {
  const statuses = getStatuses();
  const s = statuses.filter(s => s.entity_type === entityType).sort((a, b) => a.order - b.order)[0];
  return s ? s.id : '';
}

// ============================================================
// TASKS API
// ============================================================

function getTasks() {
  return sheetToObjects(getSheet('Tasks'));
}

function getTasksByJob(jobId) {
  return getTasks().filter(t => t.job_id === jobId).sort((a, b) => Number(a.order) - Number(b.order));
}

function getTasksByAssignee(assigneeId) {
  return getTasks().filter(t => t.assignee_id === assigneeId);
}

function createTask(data) {
  const sheet = getSheet('Tasks');
  const id = genId();
  const now = new Date().toISOString();
  const statusId = data.status_id || getDefaultStatusId('task');
  sheet.appendRow([
    id,
    data.job_id || '',
    data.name || '',
    data.order || 1,
    data.assignee_id || '',
    data.deadline || '',
    statusId,
    data.completed_at || '',
    data.notes || '',
    now,
  ]);
  return { id, ...data, status_id: statusId, created_at: now };
}

function updateTask(id, data) {
  const sheet = getSheet('Tasks');
  const row = findRowById(sheet, id);
  if (row === -1) return null;
  const headers = getHeaders(sheet);

  // auto set completed_at when marking hoàn thành
  if (data.status_id) {
    const statuses = getStatuses();
    const s = statuses.find(s => s.id === data.status_id);
    if (s && s.label === 'Hoàn thành' && !data.completed_at) {
      data.completed_at = new Date().toISOString();
    }
  }

  Object.keys(data).forEach(key => {
    // Không ghi đè job_id/assignee_id bằng giá trị rỗng
    if ((key === 'job_id' || key === 'assignee_id') && !data[key]) return;
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(row, col).setValue(data[key]);
  });

  // after updating task, auto-update job status
  autoUpdateJobStatus(data.job_id || getTaskById(id).job_id);

  return { id, ...data };
}

function getTaskById(id) {
  return getTasks().find(t => t.id === id) || null;
}

function deleteTask(id) {
  const sheet = getSheet('Tasks');
  const row = findRowById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
  return true;
}

function reorderTasks(jobId, orderedIds) {
  orderedIds.forEach((id, index) => {
    updateTask(id, { order: index + 1, job_id: jobId });
  });
  return true;
}

// auto update job status based on tasks
function autoUpdateJobStatus(jobId) {
  if (!jobId) return;
  const tasks = getTasksByJob(jobId);
  if (tasks.length === 0) return;

  const statuses = getStatuses();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');
  const doneStatus       = findStatusByNorm(taskStatuses, 'hoànthành') || findStatusByNorm(taskStatuses, 'đãxong');
  const inProgressStatus = findStatusByNorm(taskStatuses, 'đanglàm');

  const jobStatuses   = statuses.filter(s => s.entity_type === 'job');
  const jobDone       = findStatusByNorm(jobStatuses, 'hoànthành');
  const jobInProgress = findStatusByNorm(jobStatuses, 'đanglàm');
  const jobNotStarted = findStatusByNorm(jobStatuses, 'chưalàm');

  const allDone = tasks.every(t => doneStatus && t.status_id === doneStatus.id);
  const anyInProgress = tasks.some(t => inProgressStatus && t.status_id === inProgressStatus.id);
  const anyDone = tasks.some(t => doneStatus && t.status_id === doneStatus.id);

  let newStatusId;
  if (allDone && jobDone) newStatusId = jobDone.id;
  else if ((anyInProgress || anyDone) && jobInProgress) newStatusId = jobInProgress.id;
  else if (jobNotStarted) newStatusId = jobNotStarted.id;

  if (newStatusId) updateJob(jobId, { status_id: newStatusId });
}

// ============================================================
// TASK AVAILABILITY (sequential logic)
// ============================================================

// So sánh label không phân biệt hoa thường
function normLabel(s) {
  return (s || '').toLowerCase().replace(/\s/g, '');
}

function findStatusByNorm(statuses, norm) {
  return statuses.find(s => normLabel(s.label) === norm) || null;
}

function getEnrichedTasksForAssignee(assigneeId) {
  const allTasks = getTasks();
  const statuses = getStatuses();
  const jobs = getJobs();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  // Tìm status linh hoạt theo label
  const doneStatus    = findStatusByNorm(taskStatuses, 'hoànthành')
                     || findStatusByNorm(taskStatuses, 'đãxong');
  const ipStatus      = findStatusByNorm(taskStatuses, 'đanglàm');
  const pendingStatus = findStatusByNorm(taskStatuses, 'treo');

  // Group + sort tasks theo job
  const tasksByJob = {};
  allTasks.forEach(t => {
    if (!tasksByJob[t.job_id]) tasksByJob[t.job_id] = [];
    tasksByJob[t.job_id].push(t);
  });
  Object.values(tasksByJob).forEach(arr =>
    arr.sort((a, b) => Number(a.order) - Number(b.order))
  );

  const myTasks = allTasks.filter(t => t.assignee_id === assigneeId && t.name && String(t.name).trim());

  return myTasks.map(task => {
    const jobTasks = tasksByJob[task.job_id] || [];
    const myOrder  = Number(task.order);

    const prevTask     = jobTasks.find(t => Number(t.order) === myOrder - 1) || null;
    const prevPrevTask = jobTasks.find(t => Number(t.order) === myOrder - 2) || null;

    const isDone      = doneStatus    && task.status_id === doneStatus.id;
    const isIP        = ipStatus      && task.status_id === ipStatus.id;
    const isPending   = pendingStatus && task.status_id === pendingStatus.id;

    // Treo block task sau — không coi là done
    const isPassable  = t => doneStatus && t.status_id === doneStatus.id;
    const prevIsDone  = !prevTask    || isPassable(prevTask);
    const ppIsDone    = prevPrevTask && isPassable(prevPrevTask);

    let availability;
    if (isDone)          availability = 'done';
    else if (isPending)  availability = 'pending';
    else if (isIP)       availability = 'in_progress';
    else if (prevIsDone) availability = 'ready';
    else if (ppIsDone)   availability = 'upcoming';
    else                 availability = 'blocked';

    const job = jobs.find(j => j.id === task.job_id) || null;
    const status = taskStatuses.find(s => s.id === task.status_id) || null;
    const effectiveDeadline = task.deadline || (job ? job.deadline : '');

    return {
      ...task,
      job,
      status,
      effective_deadline: effectiveDeadline,
      availability,
      prev_task: prevTask,
    };
  });
}

// ============================================================
// DASHBOARD STATS
// ============================================================

function getDashboardStats(filters) {
  filters = filters || {};
  const dateFrom = filters.dateFrom ? new Date(filters.dateFrom) : null;
  const dateTo   = filters.dateTo   ? new Date(filters.dateTo + 'T23:59:59') : null;
  const category = filters.category || '';

  let jobs = getJobs();

  // Áp dụng filter
  if (dateFrom) jobs = jobs.filter(j => j.received_date && new Date(j.received_date) >= dateFrom);
  if (dateTo)   jobs = jobs.filter(j => j.received_date && new Date(j.received_date) <= dateTo);
  if (category) jobs = jobs.filter(j => (j.category||'') === category);

  const jobIds = new Set(jobs.map(j => j.id));
  const allTasks = getTasks();
  const tasks = allTasks.filter(t => jobIds.has(t.job_id));

  const statuses = getStatuses();
  const users = getUsers();
  const jobStatuses  = statuses.filter(s => s.entity_type === 'job');
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  const jobStatusCounts = {};
  jobStatuses.forEach(s => {
    jobStatusCounts[s.label] = jobs.filter(j => j.status_id === s.id).length;
  });

  const taskStatusCounts = {};
  taskStatuses.forEach(s => {
    taskStatusCounts[s.label] = tasks.filter(t => t.status_id === s.id).length;
  });

  const totalRevenue = jobs.reduce((sum, j) => sum + (Number(j.revenue) || 0), 0);
  const doneStatus   = findStatusByNorm(jobStatuses, 'hoànthành');
  const completedRevenue = jobs
    .filter(j => doneStatus && j.status_id === doneStatus.id)
    .reduce((sum, j) => sum + (Number(j.revenue) || 0), 0);

  const now = new Date();
  const overdueJobs = jobs.filter(j => {
    if (!j.deadline) return false;
    return (!doneStatus || j.status_id !== doneStatus.id) && new Date(j.deadline) < now;
  }).length;

  const doneTSId = findStatusByNorm(taskStatuses, 'hoànthành')?.id || findStatusByNorm(taskStatuses, 'đãxong')?.id;
  const ipTSId   = findStatusByNorm(taskStatuses, 'đanglàm')?.id;

  const employeeStats = users
    .filter(u => u.role === 'employee')
    .map(u => {
      const myTasks = tasks.filter(t => t.assignee_id === u.id);
      return {
        id: u.id, name: u.name,
        total: myTasks.length,
        done: myTasks.filter(t => t.status_id === doneTSId).length,
        in_progress: myTasks.filter(t => t.status_id === ipTSId).length,
        todo: myTasks.filter(t => t.status_id !== doneTSId && t.status_id !== ipTSId).length,
      };
    })
    .filter(e => e.total > 0);

  return {
    jobs: {
      total: jobs.length,
      by_status: jobStatusCounts,
      overdue: overdueJobs,
    },
    tasks: {
      total: tasks.length,
      by_status: taskStatusCounts,
    },
    revenue: {
      total: totalRevenue,
      completed: completedRevenue,
    },
    employees: employeeStats,
  };
}

// ============================================================
// IMAGES API — chỉ lưu avatar trên cột Jobs.avatar_id
// ============================================================

// Upload ảnh lên Drive, trả về file_id — không cần jobId
function uploadJobAvatarOnly(base64Data, mimeType, filename, uploadedBy) {
  const settings = getSettings();
  const folderSetting = settings.find(s => s.key === 'drive_folder_id');
  let folder;
  if (folderSetting && folderSetting.value) {
    try { folder = DriveApp.getFolderById(folderSetting.value); } catch(e) {}
  }
  if (!folder) {
    const folders = DriveApp.getFoldersByName('HaiLux_Images');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HaiLux_Images');
    updateSetting('drive_folder_id', folder.getId());
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { drive_file_id: file.getId() };
}

function uploadJobAvatar(base64Data, mimeType, filename, jobId, uploadedBy) {
  const settings = getSettings();
  const folderSetting = settings.find(s => s.key === 'drive_folder_id');
  let folder;
  if (folderSetting && folderSetting.value) {
    try { folder = DriveApp.getFolderById(folderSetting.value); } catch(e) {}
  }
  if (!folder) {
    const folders = DriveApp.getFoldersByName('HaiLux_Images');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HaiLux_Images');
    updateSetting('drive_folder_id', folder.getId());
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  updateJob(jobId, { avatar_id: fileId });
  return { drive_file_id: fileId, cdn_url: 'https://lh3.googleusercontent.com/d/' + fileId };
}

// ============================================================
// BULK / COMPOSITE
// ============================================================

function getJobWithDetails(jobId) {
  const job = getJob(jobId);
  if (!job) return null;
  const tasks    = getTasksByJob(jobId);
  const users    = getUsers();
  const statuses = getStatuses();
  const enrichedTasks = tasks.map(t => {
    const assignee = users.find(u => u.id === t.assignee_id);
    const status   = statuses.find(s => s.id === t.status_id);
    return { ...t, assignee, status };
  });
  const jobStatus = statuses.find(s => s.id === job.status_id);
  return { ...job, status: jobStatus, tasks: enrichedTasks };
}

function uploadTaskEvidence(base64Data, mimeType, filename, taskId, uploadedBy) {
  const settings = getSettings();
  const folderSetting = settings.find(s => s.key === 'drive_folder_id');
  let folder;
  if (folderSetting && folderSetting.value) {
    try { folder = DriveApp.getFolderById(folderSetting.value); } catch(e) {}
  }
  if (!folder) {
    const folders = DriveApp.getFoldersByName('HaiLux_Images');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HaiLux_Images');
    updateSetting('drive_folder_id', folder.getId());
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();

  // Append vào evidence_id
  const task = getTaskById(taskId);
  const existing = (task && task.evidence_id) ? String(task.evidence_id).split(',').filter(Boolean) : [];
  existing.push(fileId);
  updateTask(taskId, { evidence_id: existing.join(','), job_id: task ? task.job_id : '' });

  return { drive_file_id: fileId, cdn_url: 'https://lh3.googleusercontent.com/d/' + fileId };
}

function deleteTaskEvidence(taskId, fileId) {
  const task = getTaskById(taskId);
  if (!task) return false;
  const ids = String(task.evidence_id || '').split(',').filter(id => id && id !== fileId);
  updateTask(taskId, { evidence_id: ids.join(','), job_id: task.job_id });
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
  return true;
}

function getTaskWithImages(taskId) {
  const task = getTaskById(taskId);
  const ids = task && task.evidence_id ? String(task.evidence_id).split(',').filter(Boolean) : [];
  const images = ids.map(id => ({ drive_file_id: id, cdn_url: 'https://lh3.googleusercontent.com/d/' + id }));

  // Tính availability để block action nếu chưa đến lượt
  let availability = 'ready';
  if (task) {
    const allTasks = getTasksByJob(task.job_id).sort((a,b) => Number(a.order)-Number(b.order));
    const statuses = getStatuses().filter(s => s.entity_type === 'task');
    const doneStatus = findStatusByNorm(statuses, 'hoànthành') || findStatusByNorm(statuses, 'đãxong');
    const myOrder = Number(task.order);
    const prevTask     = allTasks.find(t => Number(t.order) === myOrder - 1) || null;
    const prevPrevTask = allTasks.find(t => Number(t.order) === myOrder - 2) || null;
    const isPassable   = t => doneStatus && t.status_id === doneStatus.id;
    const prevIsDone   = !prevTask || isPassable(prevTask);
    const ppIsDone     = prevPrevTask && isPassable(prevPrevTask);
    const ipStatus     = findStatusByNorm(statuses, 'đanglàm');
    const isDone       = doneStatus && task.status_id === doneStatus.id;
    const isIP         = ipStatus   && task.status_id === ipStatus.id;
    if (isDone)          availability = 'done';
    else if (isIP)       availability = 'in_progress';
    else if (prevIsDone) availability = 'ready';
    else if (ppIsDone)   availability = 'upcoming';
    else                 availability = 'blocked';
  }

  return { task: task, images: images, availability: availability };
}

function getAllTasksWithDetails(filters) {
  filters = filters || {};
  const tasks    = getTasks();
  const jobs     = getJobs();
  const users    = getUsers();
  const statuses = getStatuses();
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');

  const jobMap  = {};  jobs.forEach(j => jobMap[j.id] = j);
  const userMap = {};  users.forEach(u => userMap[u.id] = u);
  const stMap   = {};  taskStatuses.forEach(s => stMap[s.id] = s);

  return tasks
    .map(t => ({
      ...t,
      job:      jobMap[t.job_id]  || null,
      assignee: userMap[t.assignee_id] || null,
      status:   stMap[t.status_id]    || null,
      effective_deadline: t.deadline || (jobMap[t.job_id] ? jobMap[t.job_id].deadline : ''),
    }))
    .filter(t => {
      if (!t.name || !String(t.name).trim()) return false; // bỏ data rác
      if (filters.assignee_id && t.assignee_id !== filters.assignee_id) return false;
      if (filters.status_id   && t.status_id   !== filters.status_id)   return false;
      if (filters.deadline_from && t.effective_deadline && t.effective_deadline < filters.deadline_from) return false;
      if (filters.deadline_to   && t.effective_deadline && t.effective_deadline > filters.deadline_to + 'T23:59:59') return false;
      return true;
    })
    .sort((a, b) => {
      // Sort: deadline gần nhất lên trước, không có deadline xuống dưới
      if (!a.effective_deadline && !b.effective_deadline) return 0;
      if (!a.effective_deadline) return 1;
      if (!b.effective_deadline) return -1;
      return a.effective_deadline.localeCompare(b.effective_deadline);
    });
}

function getInitialData() {
  const jobs = getJobs();
  const categories = [...new Set(jobs.map(j => j.category).filter(Boolean))].sort();
  return {
    users: getUsers(),
    statuses: getStatuses(),
    taskTemplates: getTaskTemplates(),
    settings: getSettings(),
    categories: categories,
  };
}
