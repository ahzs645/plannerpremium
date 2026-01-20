/**
 * Import Page Script for Planner Exporter
 * Handles CSV upload, parsing, preview, and task creation
 */

document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const stepUpload = document.getElementById('step-upload');
  const stepPreview = document.getElementById('step-preview');
  const stepProgress = document.getElementById('step-progress');
  const stepResults = document.getElementById('step-results');

  const uploadZone = document.getElementById('upload-zone');
  const fileInput = document.getElementById('file-input');
  const btnBrowse = document.getElementById('btn-browse');
  const btnDownloadTemplate = document.getElementById('btn-download-template');

  const previewSummary = document.getElementById('preview-summary');
  const validationErrors = document.getElementById('validation-errors');
  const errorList = document.getElementById('error-list');
  const bucketMapList = document.getElementById('bucket-map-list');
  const previewBody = document.getElementById('preview-body');
  const previewTree = document.getElementById('preview-tree');
  const btnViewTable = document.getElementById('btn-view-table');
  const btnViewTree = document.getElementById('btn-view-tree');
  const previewTableView = document.getElementById('preview-table-view');
  const previewTreeView = document.getElementById('preview-tree-view');
  const btnBackUpload = document.getElementById('btn-back-upload');
  const btnStartImport = document.getElementById('btn-start-import');

  const progressFill = document.getElementById('progress-fill');
  const progressCurrent = document.getElementById('progress-current');
  const progressTotal = document.getElementById('progress-total');
  const progressLog = document.getElementById('progress-log');
  const btnCancelImport = document.getElementById('btn-cancel-import');

  const resultSuccess = document.getElementById('result-success');
  const resultFailed = document.getElementById('result-failed');
  const resultSkipped = document.getElementById('result-skipped');
  const failedTasks = document.getElementById('failed-tasks');
  const failedList = document.getElementById('failed-list');
  const btnImportAnother = document.getElementById('btn-import-another');
  const btnViewPlanner = document.getElementById('btn-view-planner');

  // State
  let parsedTasks = [];
  let existingBuckets = [];
  let existingResources = [];
  let bucketMapping = {}; // CSV bucket name -> existing bucket ID or null
  let importSession = null;
  let importCancelled = false;
  let plannerUrl = '';

  // CSV Template
  const CSV_TEMPLATE = `OutlineNumber,Title,Bucket,Priority,StartDate,DueDate,AssignedTo,Description,ChecklistItems
1,Phase 1: Planning,Backlog,High,2025-01-20,2025-01-31,,Project planning phase,
1.1,Define requirements,Backlog,High,2025-01-20,2025-01-22,pm@company.com,Gather requirements from stakeholders,Interview stakeholders;Document requirements;Review with team
1.2,Create project timeline,Backlog,Medium,2025-01-23,2025-01-25,pm@company.com,Build project schedule,
2,Phase 2: Development,Sprint 1,High,2025-02-01,2025-02-28,,Development phase,
2.1,Setup development environment,Sprint 1,Urgent,2025-02-01,2025-02-03,dev@company.com,Configure dev environment,Install tools;Configure CI/CD
2.1.1,Install dependencies,Sprint 1,Medium,2025-02-01,2025-02-02,dev@company.com,Install required packages,
2.2,Implement core features,Sprint 1,High,2025-02-04,2025-02-20,dev@company.com;dev2@company.com,Build main functionality,`;

  // Initialize
  init();

  function init() {
    setupEventListeners();
    fetchExistingData();
  }

  function setupEventListeners() {
    // Upload zone
    btnBrowse.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    uploadZone.addEventListener('dragover', handleDragOver);
    uploadZone.addEventListener('dragleave', handleDragLeave);
    uploadZone.addEventListener('drop', handleDrop);
    btnDownloadTemplate.addEventListener('click', downloadTemplate);

    // Preview
    btnViewTable.addEventListener('click', () => setPreviewView('table'));
    btnViewTree.addEventListener('click', () => setPreviewView('tree'));
    btnBackUpload.addEventListener('click', () => showStep('upload'));
    btnStartImport.addEventListener('click', startImport);

    // Progress
    btnCancelImport.addEventListener('click', () => { importCancelled = true; });

    // Results
    btnImportAnother.addEventListener('click', resetImport);
    btnViewPlanner.addEventListener('click', () => {
      if (plannerUrl) window.open(plannerUrl, '_blank');
    });
  }

  async function fetchExistingData() {
    try {
      const response = await chrome.runtime.sendMessage({ action: 'getImportSession' });
      if (response.success) {
        importSession = response.data;
        existingBuckets = response.data.buckets || [];
        existingResources = response.data.resources || [];
        plannerUrl = response.data.plannerUrl || '';
        console.log('[Import] Session initialized:', {
          baseUrl: importSession.baseUrl,
          buckets: existingBuckets.length,
          resources: existingResources.length
        });
      } else {
        console.error('[Import] Failed to get session:', response.error);
        showError('Failed to connect to Planner. Please navigate to a plan first and try again.');
      }
    } catch (error) {
      console.error('[Import] Error fetching session:', error);
      showError('Failed to connect to extension. Please refresh the page.');
    }
  }

  function handleDragOver(e) {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  }

  function handleDragLeave(e) {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
  }

  function handleDrop(e) {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.csv')) {
      processFile(file);
    } else {
      showError('Please drop a CSV file.');
    }
  }

  function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
      processFile(file);
    }
  }

  function processFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const csvText = e.target.result;
      parseCSV(csvText);
    };
    reader.onerror = () => {
      showError('Failed to read file.');
    };
    reader.readAsText(file);
  }

  function parseCSV(csvText) {
    const lines = csvText.split(/\r?\n/).filter(line => line.trim());
    if (lines.length < 2) {
      showError('CSV file must have a header row and at least one data row.');
      return;
    }

    const headers = parseCSVLine(lines[0]).map(h => h.toLowerCase().trim());
    const requiredHeaders = ['outlinenumber', 'title'];
    const missingHeaders = requiredHeaders.filter(h => !headers.includes(h));

    if (missingHeaders.length > 0) {
      showError(`Missing required columns: ${missingHeaders.join(', ')}`);
      return;
    }

    const tasks = [];
    const errors = [];

    for (let i = 1; i < lines.length; i++) {
      const values = parseCSVLine(lines[i]);
      const task = {};

      headers.forEach((header, index) => {
        task[header] = values[index] || '';
      });

      // Normalize field names
      const normalized = {
        rowNumber: i,
        outlineNumber: task.outlinenumber || task['outline number'] || '',
        title: task.title || '',
        bucket: task.bucket || '',
        priority: normalizePriority(task.priority || ''),
        startDate: task.startdate || task['start date'] || '',
        dueDate: task.duedate || task['due date'] || '',
        assignedTo: (task.assignedto || task['assigned to'] || '').split(';').map(e => e.trim()).filter(Boolean),
        description: task.description || task.notes || '',
        checklistItems: (task.checklistitems || task['checklist items'] || task.checklist || '').split(';').map(c => c.trim()).filter(Boolean)
      };

      // Validate
      if (!normalized.outlineNumber) {
        errors.push(`Row ${i + 1}: Missing OutlineNumber`);
      }
      if (!normalized.title) {
        errors.push(`Row ${i + 1}: Missing Title`);
      }

      // Calculate parent outline
      normalized.parentOutline = getParentOutline(normalized.outlineNumber);
      normalized.depth = normalized.outlineNumber ? normalized.outlineNumber.split('.').length - 1 : 0;

      tasks.push(normalized);
    }

    // Sort by outline number
    tasks.sort((a, b) => compareOutlineNumbers(a.outlineNumber, b.outlineNumber));

    // Validate hierarchy
    const outlineSet = new Set(tasks.map(t => t.outlineNumber));
    tasks.forEach(task => {
      if (task.parentOutline && !outlineSet.has(task.parentOutline)) {
        errors.push(`Row ${task.rowNumber + 1}: Parent outline "${task.parentOutline}" not found for "${task.outlineNumber}"`);
      }
    });

    parsedTasks = tasks;

    if (errors.length > 0) {
      showValidationErrors(errors);
    } else {
      validationErrors.classList.add('hidden');
    }

    showPreview();
  }

  function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];

      if (inQuotes) {
        if (char === '"' && nextChar === '"') {
          current += '"';
          i++;
        } else if (char === '"') {
          inQuotes = false;
        } else {
          current += char;
        }
      } else {
        if (char === '"') {
          inQuotes = true;
        } else if (char === ',') {
          result.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
    }
    result.push(current.trim());
    return result;
  }

  function normalizePriority(priority) {
    const p = priority.toLowerCase().trim();
    if (p === 'urgent' || p === '1') return 'urgent';
    if (p === 'high' || p === 'important' || p === '3') return 'high';
    if (p === 'low' || p === '9') return 'low';
    return 'medium';
  }

  function getParentOutline(outlineNumber) {
    if (!outlineNumber) return null;
    const parts = outlineNumber.split('.');
    if (parts.length === 1) return null;
    return parts.slice(0, -1).join('.');
  }

  function compareOutlineNumbers(a, b) {
    const aParts = a.split('.').map(Number);
    const bParts = b.split('.').map(Number);
    for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
      const aVal = aParts[i] || 0;
      const bVal = bParts[i] || 0;
      if (aVal !== bVal) return aVal - bVal;
    }
    return 0;
  }

  function showValidationErrors(errors) {
    errorList.innerHTML = errors.map(e => `<li>${escapeHtml(e)}</li>`).join('');
    validationErrors.classList.remove('hidden');
  }

  function showPreview() {
    showStep('preview');

    // Stats
    const rootTasks = parsedTasks.filter(t => !t.parentOutline);
    const subtasks = parsedTasks.filter(t => t.parentOutline);
    const checklistCount = parsedTasks.reduce((sum, t) => sum + t.checklistItems.length, 0);

    document.getElementById('stat-total').textContent = parsedTasks.length;
    document.getElementById('stat-root').textContent = rootTasks.length;
    document.getElementById('stat-subtasks').textContent = subtasks.length;
    document.getElementById('stat-checklists').textContent = checklistCount;

    // Bucket mapping
    const csvBuckets = [...new Set(parsedTasks.map(t => t.bucket).filter(Boolean))];
    renderBucketMapping(csvBuckets);

    // Preview table
    renderPreviewTable();
    renderPreviewTree();

    // Enable import button if no critical errors
    updateImportButton();
  }

  function renderBucketMapping(csvBuckets) {
    if (csvBuckets.length === 0) {
      document.getElementById('bucket-mapping').classList.add('hidden');
      return;
    }

    document.getElementById('bucket-mapping').classList.remove('hidden');

    bucketMapList.innerHTML = csvBuckets.map(csvBucket => {
      // Try to find matching existing bucket
      const match = existingBuckets.find(b =>
        b.name.toLowerCase() === csvBucket.toLowerCase()
      );
      const matchId = match ? match.id : '';

      return `
        <div class="mapping-row">
          <span class="mapping-csv-bucket">${escapeHtml(csvBucket)}</span>
          <span class="mapping-arrow">→</span>
          <select class="mapping-select" data-csv-bucket="${escapeHtml(csvBucket)}">
            <option value="">-- Skip (no bucket) --</option>
            ${existingBuckets.map(b => `
              <option value="${b.id}" ${b.id === matchId ? 'selected' : ''}>${escapeHtml(b.name)}</option>
            `).join('')}
          </select>
        </div>
      `;
    }).join('');

    // Update bucket mapping state
    bucketMapList.querySelectorAll('.mapping-select').forEach(select => {
      select.addEventListener('change', () => {
        updateBucketMapping();
        updateImportButton();
      });
    });

    updateBucketMapping();
  }

  function updateBucketMapping() {
    bucketMapping = {};
    bucketMapList.querySelectorAll('.mapping-select').forEach(select => {
      const csvBucket = select.dataset.csvBucket;
      bucketMapping[csvBucket] = select.value || null;
    });
  }

  function renderPreviewTable() {
    previewBody.innerHTML = parsedTasks.map(task => {
      const priorityClass = task.priority;
      const assignedCount = task.assignedTo.length;
      const checklistCount = task.checklistItems.length;

      return `
        <tr>
          <td class="outline">${escapeHtml(task.outlineNumber)}</td>
          <td style="padding-left: ${task.depth * 16}px">${escapeHtml(task.title)}</td>
          <td>${escapeHtml(task.bucket)}</td>
          <td><span class="priority ${priorityClass}">${task.priority}</span></td>
          <td>${task.dueDate || '-'}</td>
          <td>${assignedCount > 0 ? `${assignedCount} person${assignedCount > 1 ? 's' : ''}` : '-'}</td>
          <td>${checklistCount > 0 ? `${checklistCount} item${checklistCount > 1 ? 's' : ''}` : '-'}</td>
        </tr>
      `;
    }).join('');
  }

  function renderPreviewTree() {
    const tree = buildTree(parsedTasks);
    previewTree.innerHTML = tree.map(node => renderTreeNode(node)).join('');
  }

  function buildTree(tasks) {
    const taskMap = new Map();
    const roots = [];

    tasks.forEach(task => {
      taskMap.set(task.outlineNumber, { ...task, children: [] });
    });

    tasks.forEach(task => {
      const node = taskMap.get(task.outlineNumber);
      if (task.parentOutline && taskMap.has(task.parentOutline)) {
        taskMap.get(task.parentOutline).children.push(node);
      } else {
        roots.push(node);
      }
    });

    return roots;
  }

  function renderTreeNode(node) {
    const hasChildren = node.children && node.children.length > 0;
    return `
      <div class="tree-node">
        <div class="tree-node-content">
          <span class="tree-toggle">${hasChildren ? '▼' : '•'}</span>
          <span class="tree-outline">${escapeHtml(node.outlineNumber)}</span>
          <span class="tree-title ${hasChildren ? 'summary' : ''}">${escapeHtml(node.title)}</span>
        </div>
        ${hasChildren ? `
          <div class="tree-children">
            ${node.children.map(child => renderTreeNode(child)).join('')}
          </div>
        ` : ''}
      </div>
    `;
  }

  function setPreviewView(view) {
    if (view === 'table') {
      previewTableView.classList.remove('hidden');
      previewTreeView.classList.add('hidden');
      btnViewTable.classList.add('active');
      btnViewTree.classList.remove('active');
    } else {
      previewTableView.classList.add('hidden');
      previewTreeView.classList.remove('hidden');
      btnViewTable.classList.remove('active');
      btnViewTree.classList.add('active');
    }
  }

  function updateImportButton() {
    const hasErrors = !validationErrors.classList.contains('hidden');
    const hasSession = importSession && importSession.baseUrl;
    btnStartImport.disabled = hasErrors || !hasSession || parsedTasks.length === 0;
  }

  async function startImport() {
    showStep('progress');
    importCancelled = false;

    const total = parsedTasks.length;
    let created = 0;
    let failed = 0;
    let skipped = 0;
    const failedItems = [];
    const taskIdMap = {}; // outlineNumber -> created task ID

    progressTotal.textContent = total;
    progressLog.innerHTML = '';

    for (let i = 0; i < parsedTasks.length; i++) {
      if (importCancelled) {
        addLogEntry('Import cancelled by user', 'info');
        break;
      }

      const task = parsedTasks[i];
      progressCurrent.textContent = i + 1;
      progressFill.style.width = `${((i + 1) / total) * 100}%`;

      try {
        // Get bucket ID from mapping
        const bucketId = task.bucket ? bucketMapping[task.bucket] : null;

        // Get parent task ID if this is a subtask
        const parentId = task.parentOutline ? taskIdMap[task.parentOutline] : null;

        // Create task
        const response = await chrome.runtime.sendMessage({
          action: 'createImportTask',
          taskData: {
            name: task.title,
            bucketId: bucketId,
            priority: mapPriorityToValue(task.priority),
            scheduledStart: task.startDate || null,
            scheduledFinish: task.dueDate || null,
            notes: task.description || ''
          },
          baseUrl: importSession.baseUrl,
          token: importSession.token
        });

        if (response.success && response.data) {
          const createdTask = response.data;
          taskIdMap[task.outlineNumber] = createdTask.id;
          addLogEntry(`Created: ${task.title}`, 'success');

          // Set parent if subtask
          if (parentId) {
            await chrome.runtime.sendMessage({
              action: 'setTaskParent',
              taskId: createdTask.id,
              parentId: parentId,
              baseUrl: importSession.baseUrl,
              token: importSession.token
            });
          }

          // Create checklist items
          if (task.checklistItems.length > 0) {
            for (const item of task.checklistItems) {
              await chrome.runtime.sendMessage({
                action: 'createImportChecklist',
                taskId: createdTask.id,
                name: item,
                baseUrl: importSession.baseUrl,
                token: importSession.token
              });
            }
            addLogEntry(`  Added ${task.checklistItems.length} checklist items`, 'info');
          }

          created++;
        } else {
          throw new Error(response.error || 'Unknown error');
        }
      } catch (error) {
        console.error('[Import] Error creating task:', error);
        addLogEntry(`Failed: ${task.title} - ${error.message}`, 'error');
        failedItems.push({ task, error: error.message });
        failed++;
      }
    }

    showResults(created, failed, skipped, failedItems);
  }

  function mapPriorityToValue(priority) {
    switch (priority) {
      case 'urgent': return 1;
      case 'high': return 3;
      case 'low': return 9;
      default: return 5;
    }
  }

  function addLogEntry(message, type = 'info') {
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = message;
    progressLog.appendChild(entry);
    progressLog.scrollTop = progressLog.scrollHeight;
  }

  function showResults(created, failed, skipped, failedItems) {
    showStep('results');

    resultSuccess.textContent = created;
    resultFailed.textContent = failed;
    resultSkipped.textContent = skipped;

    if (failedItems.length > 0) {
      failedTasks.classList.remove('hidden');
      failedList.innerHTML = failedItems.map(item => `
        <li>
          <div class="task-title">${escapeHtml(item.task.title)}</div>
          <div class="task-error">${escapeHtml(item.error)}</div>
        </li>
      `).join('');
    } else {
      failedTasks.classList.add('hidden');
    }
  }

  function showStep(step) {
    stepUpload.classList.add('hidden');
    stepPreview.classList.add('hidden');
    stepProgress.classList.add('hidden');
    stepResults.classList.add('hidden');

    switch (step) {
      case 'upload':
        stepUpload.classList.remove('hidden');
        break;
      case 'preview':
        stepPreview.classList.remove('hidden');
        break;
      case 'progress':
        stepProgress.classList.remove('hidden');
        break;
      case 'results':
        stepResults.classList.remove('hidden');
        break;
    }
  }

  function resetImport() {
    parsedTasks = [];
    bucketMapping = {};
    fileInput.value = '';
    showStep('upload');
  }

  function showError(message) {
    alert(message); // Simple error display, can be improved
  }

  function downloadTemplate() {
    const blob = new Blob([CSV_TEMPLATE], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'planner-import-template.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
});
