/**
 * Popup Script for Planner Exporter
 */

document.addEventListener('DOMContentLoaded', async () => {
  // Elements
  const planNameEl = document.getElementById('plan-name');
  const planTypeEl = document.getElementById('plan-type');
  const tokenStatusEl = document.getElementById('token-status');
  const apiStatusEl = document.getElementById('api-status');
  const statusEl = document.getElementById('status');
  const exportBtn = document.getElementById('btn-export');
  const exportProgress = document.getElementById('export-progress');
  const addTaskBtn = document.getElementById('btn-add-task');
  const addTaskResult = document.getElementById('add-task-result');
  const taskTitleInput = document.getElementById('task-title');
  const bucketSelect = document.getElementById('bucket-select');
  const refreshBucketsBtn = document.getElementById('btn-refresh-buckets');
  const taskDueDateInput = document.getElementById('task-due-date');
  const taskPrioritySelect = document.getElementById('task-priority');
  const viewResultsBtn = document.getElementById('btn-view-results');
  const historyList = document.getElementById('history-list');

  // State
  let currentContext = null;
  let isExtracting = false;

  // Check for ongoing extraction on popup open
  async function checkExtractionState() {
    try {
      const state = await chrome.runtime.sendMessage({ action: 'getExtractionState' });

      if (state && state.status === 'extracting') {
        isExtracting = true;
        exportBtn.disabled = true;
        exportProgress.classList.remove('hidden');

        if (state.progress) {
          updateProgress(state.progress);
        }

        showStatus(`Extraction in progress (${state.method || 'API'})...`, 'info');
      } else if (state && state.status === 'complete' && state.completedAt) {
        // Show recent completion
        const completedAgo = Date.now() - state.completedAt;
        if (completedAgo < 30000) { // Within last 30 seconds
          showStatus(`Extracted ${state.taskCount} tasks!`, 'success');
        }
      } else if (state && state.status === 'error' && state.completedAt) {
        const completedAgo = Date.now() - state.completedAt;
        if (completedAgo < 30000) {
          showStatus(`Error: ${state.error}`, 'error');
        }
      }
    } catch (e) {
      console.log('Could not get extraction state:', e);
    }
  }

  // Listen for extraction state changes from background
  chrome.runtime.onMessage.addListener((message) => {
    if (message.action === 'extractionStateChanged') {
      const state = message.state;

      if (state.status === 'extracting') {
        isExtracting = true;
        exportBtn.disabled = true;
        exportProgress.classList.remove('hidden');

        if (state.progress) {
          updateProgress(state.progress);
        }
      } else if (state.status === 'complete') {
        isExtracting = false;
        exportBtn.disabled = false;
        exportProgress.classList.add('hidden');
        showStatus(`Extracted ${state.taskCount} tasks!`, 'success');
        loadHistory(); // Refresh history
      } else if (state.status === 'error') {
        isExtracting = false;
        exportBtn.disabled = false;
        exportProgress.classList.add('hidden');
        showStatus(`Error: ${state.error}`, 'error');
      }
    }

    // Also listen for progress updates
    if (message.action === 'extractionProgress' && message.progress) {
      updateProgress(message.progress);
    }
  });

  // Show status message
  function showStatus(message, type = 'info') {
    statusEl.textContent = message;
    statusEl.className = `status visible ${type}`;
    setTimeout(() => {
      statusEl.classList.remove('visible');
    }, 5000);
  }

  // Get current tab
  async function getCurrentTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    return tab;
  }

  // Check if we're on a supported page (Planner or To Do)
  function isSupportedPage(url) {
    return url && (
      url.includes('planner.cloud.microsoft') ||
      url.includes('tasks.office.com') ||
      url.includes('to-do.office.com')
    );
  }

  // Detect service type from URL
  function detectServiceType(url) {
    if (url?.includes('to-do.office.com')) return 'todo';
    if (url?.includes('tasks.office.com') || url?.includes('planner.cloud.microsoft')) return 'planner';
    return null;
  }

  // Keep backward compatibility alias
  function isPlannerPage(url) {
    return isSupportedPage(url);
  }

  // Update context display
  function updateContextDisplay(context) {
    currentContext = context;

    const isToDoService = context.serviceType === 'todo';

    // Update plan/list name
    if (isToDoService) {
      if (context.listName) {
        planNameEl.textContent = context.listName;
      } else if (context.listId) {
        planNameEl.textContent = `List: ${context.listId.substring(0, 8)}...`;
      } else {
        planNameEl.textContent = 'All To Do Lists';
      }
    } else {
      if (context.planName) {
        planNameEl.textContent = context.planName;
      } else if (context.planId) {
        planNameEl.textContent = `Plan: ${context.planId.substring(0, 8)}...`;
      } else {
        planNameEl.textContent = 'No plan detected';
      }
    }

    // Show plan/service type
    if (isToDoService) {
      planTypeEl.textContent = 'Microsoft To Do';
      planTypeEl.className = 'value';
    } else if (context.planType === 'premium') {
      planTypeEl.textContent = 'Premium (Project for the Web)';
      planTypeEl.className = 'value';
    } else if (context.planType === 'basic') {
      planTypeEl.textContent = 'Basic (Standard Planner)';
      planTypeEl.className = 'value';
    } else if (context.planId) {
      planTypeEl.textContent = 'Unknown';
      planTypeEl.className = 'value';
    } else {
      planTypeEl.textContent = '-';
      planTypeEl.className = 'value';
    }

    if (context.token) {
      tokenStatusEl.textContent = 'Found';
      tokenStatusEl.className = 'value success';
    } else {
      // For To Do, give more specific instructions
      if (isToDoService) {
        tokenStatusEl.textContent = 'Not found - scroll or click in the list';
        tokenStatusEl.className = 'value warning';
      } else {
        tokenStatusEl.textContent = 'Not found - interact with page';
        tokenStatusEl.className = 'value error';
      }
    }

    // Show API access status
    if (isToDoService && context.token) {
      apiStatusEl.textContent = 'Substrate API (To Do)';
      apiStatusEl.className = 'value success';
    } else if (context.hasPssAccess) {
      apiStatusEl.textContent = 'PSS API (Full)';
      apiStatusEl.className = 'value success';
    } else if (context.token && context.planType === 'basic') {
      apiStatusEl.textContent = 'Graph API';
      apiStatusEl.className = 'value success';
    } else if (context.planType === 'premium') {
      apiStatusEl.textContent = 'DOM only - interact with tasks';
      apiStatusEl.className = 'value';
    } else {
      apiStatusEl.textContent = '-';
      apiStatusEl.className = 'value';
    }

    // Enable/disable buttons based on context
    // For To Do, we need a token; for premium plans, we can try DOM extraction even without token
    const canExport = isToDoService
      ? !!context.token
      : (context.planType === 'premium') || (context.token && context.planId);
    exportBtn.disabled = !canExport && !context.planType && !isToDoService;

    // Add task is not supported for To Do or premium plans
    addTaskBtn.disabled = isToDoService || !context.token || context.planType === 'premium';

    // Update UI labels based on service type
    updateUILabels(isToDoService);
  }

  // Update UI labels based on service type
  function updateUILabels(isToDoService) {
    // Update the "Plan:" label to "List:" for To Do
    const planLabel = document.querySelector('#context-info .info-row:first-child .label');
    if (planLabel) {
      planLabel.textContent = isToDoService ? 'List:' : 'Plan:';
    }

    // Update export button text
    exportBtn.textContent = isToDoService ? 'Extract To Do Tasks' : 'Extract Plan Data';

    // Hide/show extraction mode options for To Do (API only, so no mode selection needed)
    const modeSelection = document.querySelector('.mode-selection');
    if (modeSelection) {
      modeSelection.style.display = isToDoService ? 'none' : 'block';
    }

    // Hide add task section for To Do
    const addTaskSection = document.getElementById('add-task-section');
    if (addTaskSection) {
      addTaskSection.style.display = isToDoService ? 'none' : 'block';
    }
  }

  // Fetch context from content script
  async function fetchContext() {
    const tab = await getCurrentTab();

    if (!isSupportedPage(tab?.url)) {
      updateContextDisplay({ token: null, planId: null, planName: null, planType: null, serviceType: null });
      showStatus('Please navigate to Microsoft Planner or To Do', 'warning');
      return;
    }

    try {
      // First, ask the page to refresh its context
      await chrome.tabs.sendMessage(tab.id, { action: 'refreshContext' }).catch(() => {});

      // Small delay to let context refresh
      await new Promise(r => setTimeout(r, 200));

      const response = await chrome.tabs.sendMessage(tab.id, { action: 'getContext' });
      console.log('[Popup] getContext response:', response);
      updateContextDisplay(response || {});
    } catch (error) {
      console.error('Failed to get context:', error);
      updateContextDisplay({ token: null, planId: null, planName: null, planType: null, serviceType: null });
      showStatus('Could not connect to page. Try refreshing.', 'error');
    }
  }

  // Load buckets
  async function loadBuckets() {
    const tab = await getCurrentTab();

    if (!currentContext?.token || !currentContext?.planId) {
      return;
    }

    bucketSelect.innerHTML = '<option value="">Loading...</option>';

    try {
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'getBuckets' });

      if (response.success && response.buckets) {
        bucketSelect.innerHTML = '<option value="">Select a bucket...</option>';
        for (const bucket of response.buckets) {
          const option = document.createElement('option');
          option.value = bucket.id;
          option.textContent = bucket.name;
          bucketSelect.appendChild(option);
        }
      } else {
        bucketSelect.innerHTML = '<option value="">Failed to load buckets</option>';
      }
    } catch (error) {
      console.error('Failed to load buckets:', error);
      bucketSelect.innerHTML = '<option value="">Error loading buckets</option>';
    }
  }

  // Get selected extraction mode
  function getSelectedMode() {
    const selected = document.querySelector('input[name="extraction-mode"]:checked');
    return selected ? selected.value : 'quick';
  }

  // Update progress display
  function updateProgress(progress) {
    const progressText = exportProgress.querySelector('.progress-text');
    const progressBar = document.getElementById('progress-bar');
    const progressDetail = document.getElementById('progress-detail');

    if (progress.message) {
      progressText.textContent = progress.message;
    }

    if (progress.total && progress.current !== undefined) {
      const percent = Math.round((progress.current / progress.total) * 100);
      progressBar.style.width = `${percent}%`;
      progressDetail.textContent = `${progress.current} of ${progress.total} tasks`;
    } else if (progress.status === 'scrolling') {
      progressBar.style.width = '0%';
      progressDetail.textContent = 'Loading all tasks...';
    }
  }

  // Export plan data
  async function exportPlan() {
    const tab = await getCurrentTab();
    const mode = getSelectedMode();

    exportBtn.disabled = true;
    exportProgress.classList.remove('hidden');

    // Reset progress
    const progressBar = document.getElementById('progress-bar');
    const progressDetail = document.getElementById('progress-detail');
    progressBar.style.width = '0%';
    progressDetail.textContent = '';

    // Set up message listener for progress updates
    const progressListener = (message, sender) => {
      if (sender.tab?.id === tab.id && message.action === 'extractionProgress') {
        updateProgress(message.progress);
      }
    };
    chrome.runtime.onMessage.addListener(progressListener);

    try {
      const response = await chrome.tabs.sendMessage(tab.id, {
        action: 'extractPlan',
        mode: mode
      });

      if (response.success) {
        showStatus('Plan exported successfully!', 'success');

        // Save to history
        await saveToHistory(response.data);

        // Open results page
        chrome.tabs.create({ url: chrome.runtime.getURL('results.html') });
      } else {
        showStatus(response.error || 'Export failed', 'error');
      }
    } catch (error) {
      console.error('Export error:', error);
      showStatus('Export failed: ' + error.message, 'error');
    } finally {
      chrome.runtime.onMessage.removeListener(progressListener);
      exportBtn.disabled = false;
      exportProgress.classList.add('hidden');
    }
  }

  // Add task
  async function addTask() {
    const tab = await getCurrentTab();
    const title = taskTitleInput.value.trim();
    const bucketId = bucketSelect.value;

    if (!title) {
      showResult(addTaskResult, 'Please enter a task title', 'error');
      return;
    }

    if (!bucketId) {
      showResult(addTaskResult, 'Please select a bucket', 'error');
      return;
    }

    addTaskBtn.disabled = true;

    const options = {};
    if (taskDueDateInput.value) {
      options.dueDateTime = new Date(taskDueDateInput.value).toISOString();
    }
    options.priority = parseInt(taskPrioritySelect.value, 10);

    try {
      const response = await chrome.tabs.sendMessage(tab.id, {
        action: 'addTask',
        bucketId: bucketId,
        title: title,
        options: options
      });

      if (response.success) {
        showResult(addTaskResult, `Task "${title}" created successfully!`, 'success');
        taskTitleInput.value = '';
        taskDueDateInput.value = '';
      } else {
        showResult(addTaskResult, response.error || 'Failed to create task', 'error');
      }
    } catch (error) {
      console.error('Add task error:', error);
      showResult(addTaskResult, 'Failed to create task: ' + error.message, 'error');
    } finally {
      addTaskBtn.disabled = false;
    }
  }

  // Show result message
  function showResult(element, message, type) {
    element.textContent = message;
    element.className = `result ${type}`;
    element.classList.remove('hidden');
    setTimeout(() => {
      element.classList.add('hidden');
    }, 5000);
  }

  // Save export to history
  async function saveToHistory(data) {
    const history = await chrome.storage.local.get('plannerExportHistory');
    const exportHistory = history.plannerExportHistory || [];

    exportHistory.unshift({
      planName: data.plan?.title || 'Unknown Plan',
      taskCount: data.tasks?.length || 0,
      exportedAt: new Date().toISOString()
    });

    // Keep only last 10 exports
    if (exportHistory.length > 10) {
      exportHistory.pop();
    }

    await chrome.storage.local.set({ plannerExportHistory: exportHistory });
    renderHistory(exportHistory);
  }

  // Render export history
  function renderHistory(history) {
    if (!history || history.length === 0) {
      historyList.innerHTML = '<p class="empty-state">No exports yet</p>';
      return;
    }

    historyList.innerHTML = history.map((item, index) => `
      <div class="history-item" data-index="${index}">
        <div class="name">${escapeHtml(item.planName)}</div>
        <div class="date">${item.taskCount} tasks - ${formatDate(item.exportedAt)}</div>
      </div>
    `).join('');
  }

  // Format date
  function formatDate(isoString) {
    const date = new Date(isoString);
    return date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  // Escape HTML
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // Load history
  async function loadHistory() {
    const history = await chrome.storage.local.get('plannerExportHistory');
    renderHistory(history.plannerExportHistory || []);
  }

  // Event listeners
  exportBtn.addEventListener('click', exportPlan);
  addTaskBtn.addEventListener('click', addTask);
  refreshBucketsBtn.addEventListener('click', loadBuckets);
  viewResultsBtn.addEventListener('click', () => {
    chrome.tabs.create({ url: chrome.runtime.getURL('results.html') });
  });

  // Import CSV button
  document.getElementById('btn-import-csv').addEventListener('click', () => {
    chrome.tabs.create({ url: chrome.runtime.getURL('import.html') });
  });

  taskTitleInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      addTask();
    }
  });

  // Initialize
  await checkExtractionState();
  await fetchContext();
  await loadBuckets();
  await loadHistory();
});
