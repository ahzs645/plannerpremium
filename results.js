/**
 * Results Page Script for Planner Exporter
 * Handles display and export of plan data
 * Supports both API data (basic plans) and DOM-scraped data (premium plans)
 */

document.addEventListener('DOMContentLoaded', async () => {
  // Elements
  const planNameEl = document.getElementById('plan-name');
  const exportDateEl = document.getElementById('export-date');
  const totalTasksEl = document.getElementById('total-tasks');
  const completedTasksEl = document.getElementById('completed-tasks');
  const inProgressTasksEl = document.getElementById('in-progress-tasks');
  const notStartedTasksEl = document.getElementById('not-started-tasks');
  const filteredCountEl = document.getElementById('filtered-count');
  const tasksListEl = document.getElementById('tasks-list');
  const bucketsListEl = document.getElementById('buckets-list');
  const filterBucketEl = document.getElementById('filter-bucket');
  const filterStatusEl = document.getElementById('filter-status');
  const searchTasksEl = document.getElementById('search-tasks');
  const btnExportJson = document.getElementById('btn-export-json');
  const btnExportCsv = document.getElementById('btn-export-csv');
  const btnExportText = document.getElementById('btn-export-text');
  const btnViewList = document.getElementById('btn-view-list');
  const btnViewHierarchy = document.getElementById('btn-view-hierarchy');
  const tasksHierarchyEl = document.getElementById('tasks-hierarchy');

  // State
  let exportData = null;
  let filteredTasks = [];
  let currentView = 'list'; // 'list' or 'hierarchy'

  // Load data from storage
  async function loadData() {
    const result = await chrome.storage.local.get('plannerExportData');
    exportData = result.plannerExportData;

    if (!exportData) {
      tasksListEl.innerHTML = '<p class="empty-state">No export data found. Please export a plan first.</p>';
      return;
    }

    renderData();
  }

  // Get bucket name for a task (handles both API and DOM data)
  function getBucketName(task) {
    // DOM-scraped data has bucketName directly
    if (task.bucketName) return task.bucketName;
    // API data has bucketId that maps to bucketMap
    if (task.bucketId && exportData.bucketMap) {
      return exportData.bucketMap[task.bucketId] || 'Unknown Bucket';
    }
    return 'No Bucket';
  }

  // Render all data
  function renderData() {
    // Header info with source indicator
    const sourceLabel = exportData.source === 'dom' ? ' (DOM)' : '';
    let planTypeLabel = '';
    if (exportData.serviceType === 'todo' || exportData.planType === 'todo') {
      planTypeLabel = ' [To Do]';
    } else if (exportData.planType === 'premium') {
      planTypeLabel = ' [Premium]';
    } else {
      planTypeLabel = ' [Basic]';
    }
    planNameEl.textContent = (exportData.planName || exportData.plan?.title || 'Unknown Plan') + planTypeLabel;
    exportDateEl.textContent = `Exported: ${formatDate(exportData.exportedAt)}${sourceLabel}`;

    // Stats
    const tasks = exportData.tasks || [];
    const completed = tasks.filter(t => getPercentComplete(t) === 100).length;
    const inProgress = tasks.filter(t => {
      const pct = getPercentComplete(t);
      return pct > 0 && pct < 100;
    }).length;
    const notStarted = tasks.filter(t => getPercentComplete(t) === 0).length;

    totalTasksEl.textContent = tasks.length;
    completedTasksEl.textContent = completed;
    inProgressTasksEl.textContent = inProgress;
    notStartedTasksEl.textContent = notStarted;

    // Build bucket list from tasks if not provided
    let buckets = exportData.buckets || [];
    if (buckets.length === 0) {
      // Extract unique buckets from tasks
      const bucketNames = new Set();
      tasks.forEach(t => {
        const name = getBucketName(t);
        if (name && name !== 'No Bucket') bucketNames.add(name);
      });
      buckets = Array.from(bucketNames).map((name, i) => ({ id: `bucket-${i}`, name }));
    }

    // Populate bucket/list filter
    const isToDoService = exportData.serviceType === 'todo' || exportData.planType === 'todo';
    filterBucketEl.innerHTML = `<option value="">${isToDoService ? 'All Lists' : 'All Buckets'}</option>`;
    for (const bucket of buckets) {
      const option = document.createElement('option');
      option.value = bucket.name; // Use name for filtering (works for both types)
      option.textContent = bucket.name;
      filterBucketEl.appendChild(option);
    }

    // Render buckets
    renderBuckets(buckets, tasks);

    // Apply filters and render tasks
    applyFilters();
  }

  // Get percent complete (handles both formats)
  function getPercentComplete(task) {
    if (typeof task.percentComplete === 'number') return task.percentComplete;
    if (typeof task.percentComplete === 'string') {
      return parseInt(task.percentComplete.replace('%', ''), 10) || 0;
    }
    return 0;
  }

  // Render buckets
  function renderBuckets(buckets, tasks) {
    if (!buckets.length) {
      bucketsListEl.innerHTML = '<p class="empty-state">No buckets found</p>';
      return;
    }

    bucketsListEl.innerHTML = buckets.map(bucket => {
      const bucketTasks = tasks.filter(t => getBucketName(t) === bucket.name);
      return `
        <div class="bucket-card">
          <div class="bucket-name">${escapeHtml(bucket.name)}</div>
          <div class="bucket-count">${bucketTasks.length} task${bucketTasks.length !== 1 ? 's' : ''}</div>
        </div>
      `;
    }).join('');
  }

  // Apply filters and render tasks
  function applyFilters() {
    const bucketFilter = filterBucketEl.value;
    const statusFilter = filterStatusEl.value;
    const searchQuery = searchTasksEl.value.toLowerCase().trim();

    let tasks = exportData.tasks || [];

    // Filter by bucket (using bucket name)
    if (bucketFilter) {
      tasks = tasks.filter(t => getBucketName(t) === bucketFilter);
    }

    // Filter by status
    if (statusFilter) {
      tasks = tasks.filter(t => {
        const status = getTaskStatus(t);
        return status === statusFilter;
      });
    }

    // Filter by search
    if (searchQuery) {
      tasks = tasks.filter(t => {
        const title = (t.title || '').toLowerCase();
        const description = (exportData.detailsMap?.[t.id]?.description || '').toLowerCase();
        const bucket = getBucketName(t).toLowerCase();
        return title.includes(searchQuery) || description.includes(searchQuery) || bucket.includes(searchQuery);
      });
    }

    filteredTasks = tasks;
    filteredCountEl.textContent = `(${tasks.length})`;
    renderTasks(tasks);

    // Also render hierarchy if that view is active
    if (currentView === 'hierarchy') {
      renderHierarchy(tasks);
    }
  }

  // Render tasks
  function renderTasks(tasks) {
    if (!tasks.length) {
      tasksListEl.innerHTML = '<p class="empty-state">No tasks match your filters</p>';
      return;
    }

    const isToDoData = exportData.serviceType === 'todo' || exportData.planType === 'todo';

    tasksListEl.innerHTML = tasks.map(task => {
      const status = getTaskStatus(task);
      const statusLabel = getStatusLabel(status);
      const bucketName = getBucketName(task);
      const details = exportData.detailsMap?.[task.id];
      const priorityClass = getPriorityClass(task.priority);
      const pct = getPercentComplete(task);

      // Handle assigned to (both formats)
      let assignedHtml = '';
      if (task.assignedTo && Array.isArray(task.assignedTo) && task.assignedTo.length > 0) {
        assignedHtml = `<div class="task-meta-item"><strong>Assigned:</strong> ${escapeHtml(task.assignedTo.join(', '))}</div>`;
      } else if (task.assignments && Object.keys(task.assignments).length > 0) {
        assignedHtml = `<div class="task-meta-item"><strong>Assigned:</strong> ${Object.keys(task.assignments).length} person(s)</div>`;
      }

      // Handle duration (DOM data)
      let durationHtml = '';
      if (task.duration) {
        durationHtml = `<div class="task-meta-item"><strong>Duration:</strong> ${escapeHtml(task.duration)}</div>`;
      }

      let checklistHtml = '';
      // Handle checklists from both detailsMap (Planner) and directly on task (To Do)
      const checklistItems = details?.checklist
        ? Object.values(details.checklist)
        : (task.checklist || []);
      if (checklistItems.length > 0) {
        checklistHtml = `
          <div class="task-checklist">
            <h4>Checklist (${checklistItems.filter(c => c.isChecked).length}/${checklistItems.length})</h4>
            ${checklistItems.map(item => `
              <div class="checklist-item ${item.isChecked ? 'completed' : ''}">
                ${item.isChecked ? '&#9745;' : '&#9744;'} ${escapeHtml(item.title)}
              </div>
            `).join('')}
          </div>
        `;
      }

      // Get description from detailsMap or directly from task
      const description = details?.description || task.description || '';

      return `
        <div class="task-card">
          <div class="task-header">
            <div class="task-title">
              <span class="priority-indicator ${priorityClass}"></span>
              ${escapeHtml(task.title)}
            </div>
            <span class="task-status ${status}">${statusLabel}${pct > 0 && pct < 100 ? ` (${pct}%)` : ''}</span>
          </div>
          <div class="task-meta">
            <div class="task-meta-item">
              <strong>${isToDoData ? 'List' : 'Bucket'}:</strong> ${escapeHtml(bucketName)}
            </div>
            ${task.startDateTime ? `
              <div class="task-meta-item">
                <strong>Start:</strong> ${formatDateSafe(task.startDateTime)}
              </div>
            ` : ''}
            ${task.dueDateTime ? `
              <div class="task-meta-item">
                <strong>Due:</strong> ${formatDateSafe(task.dueDateTime)}
              </div>
            ` : ''}
            ${durationHtml}
            ${assignedHtml}
          </div>
          ${description ? `
            <div class="task-description">${escapeHtml(description)}</div>
          ` : ''}
          ${checklistHtml}
        </div>
      `;
    }).join('');
  }

  // Get task status
  function getTaskStatus(task) {
    const pct = getPercentComplete(task);
    if (pct === 100) return 'completed';
    if (pct > 0) return 'in-progress';
    return 'not-started';
  }

  // Get status label
  function getStatusLabel(status) {
    switch (status) {
      case 'completed': return 'Completed';
      case 'in-progress': return 'In Progress';
      default: return 'Not Started';
    }
  }

  // Get priority class
  function getPriorityClass(priority) {
    switch (priority) {
      case 1: return 'priority-urgent';
      case 3: return 'priority-important';
      case 9: return 'priority-low';
      default: return 'priority-medium';
    }
  }

  // Get priority label
  function getPriorityLabel(priority) {
    switch (priority) {
      case 1: return 'Urgent';
      case 3: return 'Important';
      case 9: return 'Low';
      default: return 'Medium';
    }
  }

  // Format date (handles ISO strings)
  function formatDate(isoString) {
    if (!isoString) return '-';
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return isoString; // Return as-is if invalid
    return date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  // Format date safely (handles both ISO and display formats like "1/15/2025")
  function formatDateSafe(dateValue) {
    if (!dateValue) return '-';
    // If it's already a display format (MM/DD/YYYY), return as-is
    if (typeof dateValue === 'string' && dateValue.match(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/)) {
      return dateValue;
    }
    // Otherwise try to parse as date
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return dateValue;
    return date.toLocaleDateString();
  }

  // Escape HTML
  function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // Export to JSON
  function exportToJson() {
    const data = prepareExportData();
    const json = JSON.stringify(data, null, 2);
    downloadFile(json, `planner-export-${getFilename()}.json`, 'application/json');
  }

  // Export to CSV
  function exportToCsv() {
    const tasks = exportData.tasks || [];
    const headers = [
      'ID',
      'Title',
      'Bucket',
      'Status',
      'Priority',
      'Start Date',
      'Due Date',
      'Duration',
      'Percent Complete',
      'Assigned To',
      'Description'
    ];

    const rows = tasks.map(task => {
      const details = exportData.detailsMap?.[task.id];
      const bucketName = getBucketName(task);
      const description = details?.description || task.description || '';

      // Handle assigned to
      let assignedTo = '';
      if (task.assignedTo && Array.isArray(task.assignedTo)) {
        assignedTo = task.assignedTo.join('; ');
      } else if (task.assignments) {
        assignedTo = `${Object.keys(task.assignments).length} person(s)`;
      }

      return [
        task.id,
        csvEscape(task.title),
        csvEscape(bucketName),
        getStatusLabel(getTaskStatus(task)),
        getPriorityLabel(task.priority),
        task.startDateTime || '',
        task.dueDateTime || '',
        task.duration || '',
        getPercentComplete(task),
        csvEscape(assignedTo),
        csvEscape(description)
      ];
    });

    const csv = [headers, ...rows].map(row => row.join(',')).join('\n');
    downloadFile(csv, `planner-export-${getFilename()}.csv`, 'text/csv');
  }

  // Export to Text
  function exportToText() {
    const plan = exportData.plan || {};
    const tasks = exportData.tasks || [];
    const isToDoData = exportData.serviceType === 'todo' || exportData.planType === 'todo';

    let text = isToDoData ? `TO DO EXPORT\n` : `PLANNER EXPORT\n`;
    text += `==============\n\n`;
    text += `${isToDoData ? 'List' : 'Plan'}: ${plan.title || exportData.planName || 'Unknown'}\n`;

    let typeLabel;
    if (isToDoData) {
      typeLabel = 'Microsoft To Do';
    } else if (exportData.planType === 'premium') {
      typeLabel = 'Premium (Project for the Web)';
    } else {
      typeLabel = 'Basic (Standard Planner)';
    }
    text += `Type: ${typeLabel}\n`;
    text += `Source: ${exportData.source === 'dom' ? 'DOM Scraping' : 'Graph API'}\n`;
    text += `Exported: ${formatDate(exportData.exportedAt)}\n`;
    text += `Total Tasks: ${tasks.length}\n\n`;

    // Group tasks by bucket
    const bucketGroups = new Map();
    tasks.forEach(task => {
      const bucketName = getBucketName(task);
      if (!bucketGroups.has(bucketName)) {
        bucketGroups.set(bucketName, []);
      }
      bucketGroups.get(bucketName).push(task);
    });

    for (const [bucketName, bucketTasks] of bucketGroups) {
      text += `\n${'='.repeat(50)}\n`;
      text += `${isToDoData ? 'LIST' : 'BUCKET'}: ${bucketName}\n`;
      text += `${'='.repeat(50)}\n\n`;

      for (const task of bucketTasks) {
        const details = exportData.detailsMap?.[task.id];
        const description = details?.description || task.description || '';
        const status = getStatusLabel(getTaskStatus(task));
        const priority = getPriorityLabel(task.priority);
        const pct = getPercentComplete(task);

        text += `- [${status}${pct > 0 && pct < 100 ? ` ${pct}%` : ''}] ${task.title}\n`;
        text += `  Priority: ${priority}\n`;
        if (task.startDateTime) text += `  Start: ${formatDateSafe(task.startDateTime)}\n`;
        if (task.dueDateTime) text += `  Due: ${formatDateSafe(task.dueDateTime)}\n`;
        if (task.duration) text += `  Duration: ${task.duration}\n`;
        if (task.assignedTo?.length) text += `  Assigned: ${task.assignedTo.join(', ')}\n`;
        if (description) text += `  Description: ${description}\n`;

        // Checklist (from detailsMap or directly on task)
        const checklistItems = details?.checklist
          ? Object.values(details.checklist)
          : (task.checklist || []);
        if (checklistItems.length > 0) {
          text += `  Checklist:\n`;
          for (const item of checklistItems) {
            text += `    [${item.isChecked ? 'x' : ' '}] ${item.title}\n`;
          }
        }
        text += '\n';
      }
    }

    downloadFile(text, `planner-export-${getFilename()}.txt`, 'text/plain');
  }

  // Prepare export data
  function prepareExportData() {
    const isToDoData = exportData.serviceType === 'todo' || exportData.planType === 'todo';

    const tasks = (exportData.tasks || []).map(task => {
      const details = exportData.detailsMap?.[task.id];
      const checklistItems = details?.checklist
        ? Object.values(details.checklist)
        : (task.checklist || []);

      return {
        id: task.id,
        title: task.title,
        bucket: getBucketName(task),
        list: isToDoData ? getBucketName(task) : null,
        status: getStatusLabel(getTaskStatus(task)),
        percentComplete: getPercentComplete(task),
        priority: getPriorityLabel(task.priority),
        priorityValue: task.priority,
        startDateTime: task.startDateTime,
        dueDateTime: task.dueDateTime,
        duration: task.duration || null,
        assignedTo: task.assignedTo || (task.assignments ? Object.keys(task.assignments) : []),
        description: details?.description || task.description || null,
        checklist: checklistItems.map(c => ({
          title: c.title,
          isChecked: c.isChecked
        })),
        source: task.source || 'api'
      };
    });

    return {
      serviceType: exportData.serviceType || 'planner',
      plan: {
        id: exportData.plan?.id,
        title: exportData.plan?.title || exportData.planName,
        type: exportData.planType
      },
      buckets: (exportData.buckets || []).map(b => ({
        id: b.id,
        name: b.name
      })),
      lists: isToDoData ? (exportData.buckets || []).map(b => ({
        id: b.id,
        name: b.name
      })) : null,
      tasks: tasks,
      exportedAt: exportData.exportedAt,
      source: exportData.source,
      summary: {
        totalTasks: tasks.length,
        completed: tasks.filter(t => t.status === 'Completed').length,
        inProgress: tasks.filter(t => t.status === 'In Progress').length,
        notStarted: tasks.filter(t => t.status === 'Not Started').length
      }
    };
  }

  // CSV escape
  function csvEscape(value) {
    if (!value) return '';
    const str = String(value);
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  // Get filename
  function getFilename() {
    const planName = (exportData.planName || exportData.plan?.title || 'plan')
      .replace(/[^a-z0-9]/gi, '-')
      .toLowerCase();
    const date = new Date().toISOString().split('T')[0];
    return `${planName}-${date}`;
  }

  // Download file
  function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Build hierarchy tree from flat task list
  function buildHierarchyTree(tasks) {
    const taskMap = new Map();
    const rootTasks = [];

    // First pass: create map of all tasks
    tasks.forEach(task => {
      taskMap.set(task.id, { ...task, children: [] });
    });

    // Second pass: build tree structure
    tasks.forEach(task => {
      const taskNode = taskMap.get(task.id);
      if (task.parentId && taskMap.has(task.parentId)) {
        taskMap.get(task.parentId).children.push(taskNode);
      } else {
        rootTasks.push(taskNode);
      }
    });

    // Sort children by order
    const sortChildren = (nodes) => {
      nodes.sort((a, b) => (a.order || 0) - (b.order || 0));
      nodes.forEach(node => {
        if (node.children.length > 0) {
          sortChildren(node.children);
        }
      });
    };
    sortChildren(rootTasks);

    return rootTasks;
  }

  // Render hierarchy view
  function renderHierarchy(tasks) {
    if (!tasks.length) {
      tasksHierarchyEl.innerHTML = '<p class="empty-state">No tasks match your filters</p>';
      return;
    }

    const tree = buildHierarchyTree(tasks);

    function renderNode(node, depth = 0) {
      const status = getTaskStatus(node);
      const statusLabel = getStatusLabel(status);
      const priorityClass = getPriorityClass(node.priority);
      const pct = getPercentComplete(node);
      const indent = depth * 24;
      const hasChildren = node.children && node.children.length > 0;
      const isSummary = node.isSummaryTask;

      // Assigned info
      let assignedText = '';
      if (node.assignedTo && node.assignedTo.length > 0) {
        assignedText = node.assignedTo.slice(0, 2).join(', ');
        if (node.assignedTo.length > 2) assignedText += ` +${node.assignedTo.length - 2}`;
      }

      let html = `
        <div class="hierarchy-item ${isSummary ? 'summary-task' : ''}" style="padding-left: ${indent}px;">
          <div class="hierarchy-row">
            <span class="hierarchy-toggle">${hasChildren ? '▼' : '•'}</span>
            <span class="hierarchy-number">${escapeHtml(node.outlineNumber || '')}</span>
            <span class="priority-indicator ${priorityClass}"></span>
            <span class="hierarchy-title ${isSummary ? 'bold' : ''}">${escapeHtml(node.title)}</span>
            <span class="hierarchy-status ${status}">${statusLabel}${pct > 0 && pct < 100 ? ` ${pct}%` : ''}</span>
            ${node.dueDateTime ? `<span class="hierarchy-date">${formatDateSafe(node.dueDateTime)}</span>` : ''}
            ${assignedText ? `<span class="hierarchy-assigned">${escapeHtml(assignedText)}</span>` : ''}
          </div>
        </div>
      `;

      if (hasChildren) {
        node.children.forEach(child => {
          html += renderNode(child, depth + 1);
        });
      }

      return html;
    }

    tasksHierarchyEl.innerHTML = tree.map(node => renderNode(node, 0)).join('');
  }

  // Toggle view
  function setView(view) {
    currentView = view;

    if (view === 'list') {
      tasksListEl.classList.remove('hidden');
      tasksHierarchyEl.classList.add('hidden');
      btnViewList.classList.add('active');
      btnViewHierarchy.classList.remove('active');
    } else {
      tasksListEl.classList.add('hidden');
      tasksHierarchyEl.classList.remove('hidden');
      btnViewList.classList.remove('active');
      btnViewHierarchy.classList.add('active');
    }

    // Re-render with current filters
    if (view === 'hierarchy') {
      renderHierarchy(filteredTasks);
    }
  }

  // Event listeners
  filterBucketEl.addEventListener('change', applyFilters);
  filterStatusEl.addEventListener('change', applyFilters);
  searchTasksEl.addEventListener('input', applyFilters);
  btnExportJson.addEventListener('click', exportToJson);
  btnExportCsv.addEventListener('click', exportToCsv);
  btnExportText.addEventListener('click', exportToText);
  btnViewList.addEventListener('click', () => setView('list'));
  btnViewHierarchy.addEventListener('click', () => setView('hierarchy'));

  // Initialize
  await loadData();
});
