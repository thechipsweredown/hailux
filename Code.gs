// ============================================================
// ENTRY POINTS
// ============================================================

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
    Jobs:          ['id','code','name','category','customer_name','customer_contact','received_date','deadline','revenue','repair_scope','status_id','notes','created_at'],
    Tasks:         ['id','job_id','name','order','assignee_id','deadline','status_id','completed_at','notes','created_at'],
    Images:        ['id','entity_type','entity_id','drive_file_id','drive_url','uploaded_by','uploaded_at','caption'],
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
  range.setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
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
    ];
    statuses.forEach(r => statusSheet.appendRow(r));
  }

  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['manager_password', '1234']);
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
  if (!ss) throw new Error('Không tìm thấy Spreadsheet. Script cần được gắn vào Google Sheet.');
  return ss.getSheetByName(name);
}

function sheetToObjects(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
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

function verifyManagerPassword(password) {
  const settings = sheetToObjects(getSheet('Settings'));
  const entry = settings.find(s => s.key === 'manager_password');
  return entry && String(entry.value) === String(password);
}

function debugAll() {
  var results = {};
  var sheets = ['Users','Jobs','Tasks','Statuses','TaskTemplates','Settings','Images'];
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
  deleteImagesForEntity('job', id);
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
  deleteImagesForEntity('task', id);
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
  const doneStatus = taskStatuses.find(s => s.label === 'Hoàn thành');
  const inProgressStatus = taskStatuses.find(s => s.label === 'Đang làm');

  const jobStatuses = statuses.filter(s => s.entity_type === 'job');
  const jobDone = jobStatuses.find(s => s.label === 'Hoàn thành');
  const jobInProgress = jobStatuses.find(s => s.label === 'Đang làm');
  const jobNotStarted = jobStatuses.find(s => s.label === 'Chưa làm');

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

function getEnrichedTasksForAssignee(assigneeId) {
  const allTasks = getTasks();
  const statuses = getStatuses();
  const users = getUsers();
  const jobs = getJobs();
  const jobStatuses = statuses.filter(s => s.entity_type === 'job');
  const taskStatuses = statuses.filter(s => s.entity_type === 'task');
  const doneStatus = taskStatuses.find(s => s.label === 'Hoàn thành');

  // group tasks by job
  const tasksByJob = {};
  allTasks.forEach(t => {
    if (!tasksByJob[t.job_id]) tasksByJob[t.job_id] = [];
    tasksByJob[t.job_id].push(t);
  });
  Object.values(tasksByJob).forEach(arr => arr.sort((a, b) => Number(a.order) - Number(b.order)));

  const myTasks = allTasks.filter(t => t.assignee_id === assigneeId);

  return myTasks.map(task => {
    const jobTasks = (tasksByJob[task.job_id] || []);
    const myOrder = Number(task.order);
    const prevTask = jobTasks.find(t => Number(t.order) === myOrder - 1);
    const prevDone = !prevTask || (doneStatus && prevTask.status_id === doneStatus.id);
    const prevPrevTask = jobTasks.find(t => Number(t.order) === myOrder - 2);
    const prevPrevDone = !prevPrevTask || (doneStatus && prevPrevTask.status_id === doneStatus.id);

    const job = jobs.find(j => j.id === task.job_id);
    const status = taskStatuses.find(s => s.id === task.status_id);
    const effectiveDeadline = task.deadline || (job ? job.deadline : '');

    let availability = 'blocked'; // chưa đến lượt
    if (doneStatus && task.status_id === doneStatus.id) {
      availability = 'done';
    } else if (prevDone) {
      // task trước done → available
      availability = prevPrevDone ? 'ready' : 'ready'; // có thể làm
      // "sắp phải làm" = prevTask tồn tại và prevPrevDone (task trước task của mình vừa done)
      if (prevTask && doneStatus && prevTask.status_id === doneStatus.id) {
        availability = 'upcoming'; // sắp phải làm
      }
      if (!prevTask) availability = 'ready'; // task đầu tiên luôn available
    }

    // override: nếu đang làm thì là 'in_progress'
    const inProgressStatus = taskStatuses.find(s => s.label === 'Đang làm');
    if (inProgressStatus && task.status_id === inProgressStatus.id) availability = 'in_progress';

    return {
      ...task,
      job,
      status,
      effective_deadline: effectiveDeadline,
      availability,
      prev_task: prevTask || null,
    };
  });
}

// ============================================================
// DASHBOARD STATS
// ============================================================

function getDashboardStats() {
  const jobs = getJobs();
  const tasks = getTasks();
  const statuses = getStatuses();
  const users = getUsers();

  const jobStatuses = statuses.filter(s => s.entity_type === 'job');
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

  const doneStatus = jobStatuses.find(s => s.label === 'Hoàn thành');
  const completedRevenue = jobs
    .filter(j => doneStatus && j.status_id === doneStatus.id)
    .reduce((sum, j) => sum + (Number(j.revenue) || 0), 0);

  const now = new Date();
  const overdueJobs = jobs.filter(j => {
    if (!j.deadline) return false;
    const doneId = doneStatus ? doneStatus.id : null;
    return j.status_id !== doneId && new Date(j.deadline) < now;
  }).length;

  // per-employee task counts
  const employeeStats = users
    .filter(u => u.role === 'employee')
    .map(u => {
      const myTasks = tasks.filter(t => t.assignee_id === u.id);
      const doneSId = taskStatuses.find(s => s.label === 'Hoàn thành')?.id;
      const inProgSId = taskStatuses.find(s => s.label === 'Đang làm')?.id;
      return {
        id: u.id,
        name: u.name,
        total: myTasks.length,
        done: myTasks.filter(t => t.status_id === doneSId).length,
        in_progress: myTasks.filter(t => t.status_id === inProgSId).length,
        todo: myTasks.filter(t => t.status_id !== doneSId && t.status_id !== inProgSId).length,
      };
    });

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
// IMAGES API
// ============================================================

function getImages(entityType, entityId) {
  return sheetToObjects(getSheet('Images'))
    .filter(img => img.entity_type === entityType && img.entity_id === entityId);
}

function uploadImage(base64Data, mimeType, filename, entityType, entityId, uploadedBy, caption) {
  const settings = getSettings();
  const folderSetting = settings.find(s => s.key === 'drive_folder_id');

  let folder;
  if (folderSetting && folderSetting.value) {
    folder = DriveApp.getFolderById(folderSetting.value);
  } else {
    // create default folder
    const folders = DriveApp.getFoldersByName('HaiLux_Images');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HaiLux_Images');
    updateSetting('drive_folder_id', folder.getId());
  }

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const driveUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

  const sheet = getSheet('Images');
  const id = genId();
  const now = new Date().toISOString();
  sheet.appendRow([id, entityType, entityId, file.getId(), driveUrl, uploadedBy, now, caption || '']);

  return { id, entity_type: entityType, entity_id: entityId, drive_file_id: file.getId(), drive_url: driveUrl, uploaded_by: uploadedBy, uploaded_at: now, caption };
}

function deleteImage(id) {
  const images = sheetToObjects(getSheet('Images'));
  const img = images.find(i => i.id === id);
  if (img) {
    try { DriveApp.getFileById(img.drive_file_id).setTrashed(true); } catch(e) {}
    const sheet = getSheet('Images');
    const row = findRowById(sheet, id);
    if (row !== -1) sheet.deleteRow(row);
  }
  return true;
}

function deleteImagesForEntity(entityType, entityId) {
  const images = sheetToObjects(getSheet('Images'))
    .filter(i => i.entity_type === entityType && i.entity_id === entityId);
  images.forEach(i => deleteImage(i.id));
}

// ============================================================
// BULK / COMPOSITE
// ============================================================

function getJobWithDetails(jobId) {
  const job = getJob(jobId);
  if (!job) return null;
  const tasks = getTasksByJob(jobId);
  const images = getImages('job', jobId);
  const users = getUsers();
  const statuses = getStatuses();

  const enrichedTasks = tasks.map(t => {
    const taskImages = getImages('task', t.id);
    const assignee = users.find(u => u.id === t.assignee_id);
    const status = statuses.find(s => s.id === t.status_id);
    return { ...t, assignee, status, images: taskImages };
  });

  const jobStatus = statuses.find(s => s.id === job.status_id);
  return { ...job, status: jobStatus, tasks: enrichedTasks, images };
}

function getTaskWithImages(taskId) {
  var task = getTaskById(taskId);
  var images = getImages('task', taskId);
  return { task: task, images: images };
}

function getInitialData() {
  return {
    users: getUsers(),
    statuses: getStatuses(),
    taskTemplates: getTaskTemplates(),
    settings: getSettings(),
  };
}
