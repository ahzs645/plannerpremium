/**
 * Popup script for Microsoft Planner Interface
 */

class PlannerPopup {
  constructor() {
    this.currentData = null;
    this.isConnected = false;
    this.init();
  }

  async init() {
    this.setupEventListeners();
    await this.checkConnection();
    await this.loadCurrentData();
    this.updateUI();
  }

  setupEventListeners() {
    // Extract data button
    document.getElementById('extractDataBtn').addEventListener('click', () => {
      this.extractData();
    });

    // Export buttons
    document.getElementById('exportDataBtn').addEventListener('click', () => {
      this.toggleExportOptions();
    });

    document.getElementById('copyDataBtn').addEventListener('click', () => {
      this.copyToClipboard();
    });

    // Export format buttons
    document.getElementById('exportJson').addEventListener('click', () => {
      this.exportData('json');
    });

    document.getElementById('exportCsv').addEventListener('click', () => {
      this.exportData('csv');
    });

    document.getElementById('exportTxt').addEventListener('click', () => {
      this.exportData('txt');
    });

    // Show all tasks button
    document.getElementById('showAllTasksBtn').addEventListener('click', () => {
      this.showAllTasks();
    });

    // Retry button
    document.getElementById('retryBtn').addEventListener('click', () => {
      this.init();
    });

    // Auto-extract checkbox
    document.getElementById('autoExtract').addEventListener('change', (e) => {
      this.saveSettings({ autoExtract: e.target.checked });
    });

    // Real-time updates checkbox
    document.getElementById('realTimeUpdates').addEventListener('change', (e) => {
      this.saveSettings({ realTimeUpdates: e.target.checked });
    });
  }

  async checkConnection() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (!tab || !tab.url) {
        this.setConnectionStatus(false, 'No active tab');
        return;
      }

      // Check if we're on a Planner page
      const isPlannerPage = tab.url.includes('planner.cloud.microsoft') ||
                           tab.url.includes('tasks.office.com');

      if (!isPlannerPage) {
        this.setConnectionStatus(false, 'Not on Planner page');
        return;
      }

      // Try to communicate with content script
      try {
        const response = await chrome.tabs.sendMessage(tab.id, { action: 'getCurrentData' });
        this.setConnectionStatus(true, 'Connected to Planner');
        this.isConnected = true;
      } catch (error) {
        this.setConnectionStatus(false, 'Content script not loaded');
      }
    } catch (error) {
      this.setConnectionStatus(false, 'Connection error');
    }
  }

  setConnectionStatus(connected, message) {
    const statusDot = document.getElementById('statusDot');
    const statusText = document.getElementById('statusText');

    statusDot.className = `status-dot ${connected ? 'connected' : 'error'}`;
    statusText.textContent = message;
    this.isConnected = connected;
  }

  async loadCurrentData() {
    try {
      // Try to get data from session storage first
      const sessionData = sessionStorage.getItem('currentPlannerData');
      if (sessionData) {
        this.currentData = this.sanitizeData(JSON.parse(sessionData));
        return;
      }

      // Try to get data from content script
      if (this.isConnected) {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        const response = await chrome.tabs.sendMessage(tab.id, { action: 'getCurrentData' });

        if (response && response.success) {
          this.currentData = this.sanitizeData(response.data);
        }
      }

      // Fallback to stored data
      const result = await chrome.storage.local.get(['plannerData']);
      if (result.plannerData) {
        this.currentData = this.sanitizeData(result.plannerData);
      }
    } catch (error) {
      console.error('Error loading data:', error);
      this.showError('Failed to load data');
    }
  }

  async extractData() {
    if (!this.isConnected) {
      this.showError('Not connected to Planner page');
      return;
    }

    try {
      this.setButtonLoading('extractDataBtn', true);

      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'extractPlannerData' });

      if (response && response.success) {
        this.currentData = this.sanitizeData(response.data);
        this.updateUI();
        this.showSuccess('Data extracted successfully');

        // Store the data
        await chrome.storage.local.set({ plannerData: this.currentData });
      } else {
        this.showError('Failed to extract data');
      }
    } catch (error) {
      console.error('Error extracting data:', error);
      this.showError('Error extracting data');
    } finally {
      this.setButtonLoading('extractDataBtn', false);
    }
  }

  updateUI() {
    this.hideError();

    this.currentData = this.sanitizeData(this.currentData);

    if (!this.currentData) {
      this.showNoData();
      return;
    }

    const planData = this.currentData.planData || {};
    const taskData = Array.isArray(this.currentData.taskData)
      ? this.currentData.taskData
      : [];

    // Update plan information
    document.getElementById('planName').textContent = planData.planName || 'Unknown';
    document.getElementById('currentView').textContent = planData.currentView || 'Unknown';
    document.getElementById('accessLevel').textContent = planData.accessLevel || 'Unknown';
    document.getElementById('bucketsCount').textContent =
      Array.isArray(planData.buckets) ? planData.buckets.length : 0;

    // Update task summary
    document.getElementById('totalTasks').textContent = taskData.length;

    const completed = taskData.filter(task => task.completed || task.progress === 100).length;
    const inProgress = taskData.filter(task =>
      !task.completed && task.progress > 0 && task.progress < 100
    ).length;
    const notStarted = taskData.filter(task =>
      !task.completed && (!task.progress || task.progress === 0)
    ).length;

    document.getElementById('completedTasks').textContent = completed;
    document.getElementById('inProgressTasks').textContent = inProgress;
    document.getElementById('notStartedTasks').textContent = notStarted;

    // Show task list if we have tasks
    if (taskData.length > 0) {
      this.populateTaskList(taskData.slice(0, 5)); // Show first 5 tasks
      document.getElementById('taskListSection').style.display = 'block';
    } else {
      document.getElementById('taskListSection').style.display = 'none';
    }

    // Update last updated time
    const timestamp = this.currentData.timestamp || new Date().toISOString();
    document.getElementById('lastUpdated').textContent =
      new Date(timestamp).toLocaleTimeString();
  }

  populateTaskList(tasks) {
    if (!Array.isArray(tasks) || tasks.length === 0) {
      document.getElementById('taskListSection').style.display = 'none';
      return;
    }

    const taskList = document.getElementById('taskList');
    taskList.innerHTML = '';

    tasks.forEach(task => {
      const taskElement = document.createElement('div');
      taskElement.className = 'task-item';

      const cleanedName = this.cleanTaskName(task.name);
      const taskName = cleanedName || 'Untitled Task';

      taskElement.innerHTML = `
        <div class="task-name">${taskName}</div>
        <div class="task-meta">
          <span>${task.assignedTo || 'Unassigned'}</span>
          <span>${task.progress || 0}% complete</span>
        </div>
      `;

      taskList.appendChild(taskElement);
    });
  }

  sanitizeData(data) {
    if (!data) return null;

    const planData = { ...(data.planData || {}) };
    const taskData = this.getDisplayableTasks(data.taskData || []);

    return {
      ...data,
      planData,
      taskData
    };
  }

  getDisplayableTasks(taskData) {
    if (!Array.isArray(taskData)) return [];

    const seen = new Set();
    const result = [];

    taskData.forEach(task => {
      if (!this.isMeaningfulTask(task)) {
        return;
      }

      const cleanedName = this.cleanTaskName(task.name);
      const key = (task.id || cleanedName).toLowerCase();

      if (!key || seen.has(key)) {
        return;
      }

      seen.add(key);
      result.push({
        ...task,
        name: cleanedName
      });
    });

    return result;
  }

  isMeaningfulTask(task) {
    if (!task) return false;

    const cleanedName = this.cleanTaskName(task.name);
    if (!cleanedName) return false;

    const normalized = cleanedName.toLowerCase();
    const simplified = normalized.replace(/[^a-z0-9\s]/g, '');

    const exactMatches = new Set([
      'add new task',
      'filters',
      'grid',
      'board',
      'reports',
      'charts',
      'my plans',
      'assigned to',
      'start',
      'finish',
      'actions',
      'automation',
      'extract data',
      'export json',
      'copy to clipboard',
      'quick look',
      'unassigned'
    ]);

    if (exactMatches.has(normalized) || exactMatches.has(simplified)) {
      return false;
    }

    const prefixMatches = [
      'basic access',
      'show all tasks',
      'connected to planner',
      'plan information',
      'plan name',
      'current view',
      'total tasks',
      'tasks summary',
      'recent tasks',
      'last updated'
    ];

    if (prefixMatches.some(entry =>
      this.startsWithTerm(normalized, entry) || this.startsWithTerm(simplified, entry)
    )) {
      return false;
    }

    if (/^\d+%$/.test(normalized)) {
      return false;
    }

    return true;
  }

  cleanTaskName(rawName) {
    if (!rawName) return '';

    let taskName = String(rawName);

    if (taskName.includes('Task Name ')) {
      const match = taskName.match(/Task Name\s+([^.]+)/);
      if (match) {
        taskName = match[1];
      }
    }

    taskName = taskName
      .replace(/Use the space or enter key.*$/i, '')
      .replace(/\s+/g, ' ')
      .replace(/\.+\s*$/, '')
      .trim();

    return taskName;
  }

  startsWithTerm(value, term) {
    if (!value || !term) return false;
    if (!value.startsWith(term)) return false;
    const nextChar = value.charAt(term.length);
    if (!nextChar) {
      return true;
    }
    return /[^a-z0-9]/i.test(nextChar);
  }

  showAllTasks() {
    if (!this.currentData || !this.currentData.taskData) return;

    // Create a new window/tab with all tasks
    const allTasksData = {
      planData: this.currentData.planData,
      taskData: this.currentData.taskData,
      timestamp: this.currentData.timestamp
    };

    const dataUrl = 'data:text/html;charset=utf-8,' + encodeURIComponent(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>All Planner Tasks</title>
        <style>
          body { font-family: 'Segoe UI', sans-serif; margin: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #6264a7; color: white; }
          tr:nth-child(even) { background-color: #f2f2f2; }
          h1 { color: #6264a7; }
        </style>
      </head>
      <body>
        <h1>All Tasks - ${allTasksData.planData?.planName || 'Unknown Plan'}</h1>
        <p>Extracted: ${new Date(allTasksData.timestamp).toLocaleString()}</p>
        <p>Access Level: ${allTasksData.planData?.accessLevel || 'Unknown'}</p>
        <table>
          <thead>
            <tr>
              <th>Task Name</th>
              <th>Assigned To</th>
              <th>Progress</th>
              <th>Priority</th>
              <th>Bucket</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${allTasksData.taskData.map(task => `
              <tr>
                <td>${task.name || 'Untitled'}</td>
                <td>${task.assignedTo || 'Unassigned'}</td>
                <td>${task.progress || 0}%</td>
                <td>${task.priority || 'Normal'}</td>
                <td>${task.bucket || 'No Bucket'}</td>
                <td>${task.completed ? 'Completed' : 'Active'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </body>
      </html>
    `);

    chrome.tabs.create({ url: dataUrl });
  }

  toggleExportOptions() {
    const exportSection = document.getElementById('exportOptionsSection');
    exportSection.style.display = exportSection.style.display === 'none' ? 'block' : 'none';
  }

  async copyToClipboard() {
    if (!this.currentData) {
      this.showError('No data to copy');
      return;
    }

    try {
      const jsonData = JSON.stringify(this.currentData, null, 2);
      await navigator.clipboard.writeText(jsonData);
      this.showSuccess('Data copied to clipboard');
    } catch (error) {
      console.error('Error copying to clipboard:', error);
      this.showError('Failed to copy data');
    }
  }

  exportData(format) {
    if (!this.currentData) {
      this.showError('No data to export');
      return;
    }

    const filename = `planner-data-${new Date().toISOString().split('T')[0]}`;
    let content, mimeType;

    switch (format) {
      case 'json':
        content = JSON.stringify(this.currentData, null, 2);
        mimeType = 'application/json';
        break;

      case 'csv':
        content = this.convertToCSV(this.currentData.taskData);
        mimeType = 'text/csv';
        break;

      case 'txt':
        content = this.convertToText(this.currentData);
        mimeType = 'text/plain';
        break;

      default:
        this.showError('Unknown export format');
        return;
    }

    this.downloadFile(content, `${filename}.${format}`, mimeType);
    this.showSuccess(`Data exported as ${format.toUpperCase()}`);
  }

  convertToCSV(tasks) {
    if (!tasks || !Array.isArray(tasks)) return '';

    const headers = ['Name', 'Assigned To', 'Progress', 'Priority', 'Bucket', 'Completed'];
    const rows = tasks.map(task => [
      task.name || '',
      task.assignedTo || '',
      task.progress || 0,
      task.priority || '',
      task.bucket || '',
      task.completed ? 'Yes' : 'No'
    ]);

    return [headers, ...rows]
      .map(row => row.map(field => `"${field}"`).join(','))
      .join('\n');
  }

  convertToText(data) {
    let text = `Microsoft Planner Data Export\n`;
    text += `Generated: ${new Date().toLocaleString()}\n\n`;

    if (data.planData) {
      text += `Plan Information:\n`;
      text += `- Name: ${data.planData.planName || 'Unknown'}\n`;
      text += `- Current View: ${data.planData.currentView || 'Unknown'}\n`;
      text += `- Access Level: ${data.planData.accessLevel || 'Unknown'}\n`;
      text += `- Buckets: ${data.planData.buckets?.length || 0}\n\n`;
    }

    if (data.taskData && Array.isArray(data.taskData)) {
      text += `Tasks (${data.taskData.length}):\n`;
      data.taskData.forEach((task, index) => {
        text += `${index + 1}. ${task.name || 'Untitled'}\n`;
        text += `   Assigned: ${task.assignedTo || 'Unassigned'}\n`;
        text += `   Progress: ${task.progress || 0}%\n`;
        text += `   Status: ${task.completed ? 'Completed' : 'Active'}\n\n`;
      });
    }

    return text;
  }

  downloadFile(content, filename, mimeType) {
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

  async saveSettings(settings) {
    try {
      const current = await chrome.storage.local.get(['settings']);
      const updatedSettings = { ...current.settings, ...settings };
      await chrome.storage.local.set({ settings: updatedSettings });
    } catch (error) {
      console.error('Error saving settings:', error);
    }
  }

  setButtonLoading(buttonId, loading) {
    const button = document.getElementById(buttonId);
    if (loading) {
      button.disabled = true;
      button.innerHTML = '<span class="loading"></span> Loading...';
    } else {
      button.disabled = false;
      button.innerHTML = button.dataset.originalText || button.innerHTML;
    }
  }

  showError(message) {
    document.getElementById('errorSection').style.display = 'block';
    document.getElementById('errorMessage').textContent = message;
  }

  hideError() {
    document.getElementById('errorSection').style.display = 'none';
  }

  showSuccess(message) {
    // You could implement a success notification here
    console.log('Success:', message);
  }

  showNoData() {
    // Show a message when no data is available
    const sections = ['planInfoSection', 'tasksSummarySection'];
    sections.forEach(sectionId => {
      const section = document.getElementById(sectionId);
      if (section) {
        section.style.opacity = '0.5';
      }
    });
  }
}

// Initialize the popup when the DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
  new PlannerPopup();
});
