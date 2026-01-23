/**
 * Background Service Worker for Planner Exporter
 * Handles API calls, token storage, and file downloads
 */

// ============================================
// API CONFIGURATION
// ============================================

const GRAPH_API = 'https://graph.microsoft.com/v1.0';
const PSS_API = 'https://project.microsoft.com/pss/api/v1.0';

// ============================================
// EXTRACTION STATE MANAGEMENT
// ============================================

// State: idle, extracting, complete, error
let extractionState = {
  status: 'idle',
  progress: null,
  startedAt: null,
  completedAt: null,
  error: null,
  taskCount: 0,
  method: null
};

async function updateExtractionState(newState) {
  extractionState = { ...extractionState, ...newState };

  // Persist to storage
  await chrome.storage.local.set({ extractionState });

  // Update badge
  updateBadge(extractionState);

  // Broadcast to any open popups
  chrome.runtime.sendMessage({
    action: 'extractionStateChanged',
    state: extractionState
  }).catch(() => {}); // Ignore if no listeners
}

function updateBadge(state) {
  if (state.status === 'extracting') {
    const progress = state.progress;
    if (progress?.total && progress?.current !== undefined) {
      const pct = Math.round((progress.current / progress.total) * 100);
      chrome.action.setBadgeText({ text: `${pct}%` });
      chrome.action.setBadgeBackgroundColor({ color: '#0078d4' });
    } else {
      chrome.action.setBadgeText({ text: '...' });
      chrome.action.setBadgeBackgroundColor({ color: '#0078d4' });
    }
  } else if (state.status === 'complete') {
    chrome.action.setBadgeText({ text: '✓' });
    chrome.action.setBadgeBackgroundColor({ color: '#107c10' });
    // Clear badge after 5 seconds
    setTimeout(() => {
      chrome.action.setBadgeText({ text: '' });
    }, 5000);
  } else if (state.status === 'error') {
    chrome.action.setBadgeText({ text: '!' });
    chrome.action.setBadgeBackgroundColor({ color: '#d13438' });
    // Clear badge after 5 seconds
    setTimeout(() => {
      chrome.action.setBadgeText({ text: '' });
    }, 5000);
  } else {
    chrome.action.setBadgeText({ text: '' });
  }
}

// Restore state on startup
chrome.storage.local.get('extractionState', (result) => {
  if (result.extractionState) {
    extractionState = result.extractionState;
    // If it was extracting when service worker died, mark as error
    if (extractionState.status === 'extracting') {
      extractionState.status = 'error';
      extractionState.error = 'Extraction interrupted';
    }
  }
});

// ============================================
// TOKEN STORAGE
// ============================================

const TOKEN_KEYS = {
  GRAPH: 'plannerGraphToken',
  PSS: 'plannerPssToken',
  PSS_PROJECT: 'plannerPssProjectId',
  TODO: 'todoToken',
  TODO_LIST: 'todoListId'
};

async function storeToken(type, token, metadata = {}) {
  const key = TOKEN_KEYS[type];
  if (!key) return;

  await chrome.storage.local.set({
    [key]: {
      token,
      capturedAt: Date.now(),
      ...metadata
    }
  });
  console.log(`[Background] Stored ${type} token`);
}

async function getToken(type) {
  const key = TOKEN_KEYS[type];
  if (!key) return null;

  const result = await chrome.storage.local.get(key);
  return result[key];
}

async function getAllTokens() {
  const result = await chrome.storage.local.get([
    TOKEN_KEYS.GRAPH,
    TOKEN_KEYS.PSS,
    TOKEN_KEYS.PSS_PROJECT
  ]);

  return {
    graphToken: result[TOKEN_KEYS.GRAPH]?.token || null,
    pssToken: result[TOKEN_KEYS.PSS]?.token || null,
    pssProjectId: result[TOKEN_KEYS.PSS_PROJECT]?.projectId || null
  };
}

// ============================================
// API WRAPPERS
// ============================================

async function graphFetch(endpoint, token, options = {}) {
  const url = endpoint.startsWith('http') ? endpoint : `${GRAPH_API}${endpoint}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers
    }
  });

  return response;
}

async function pssFetch(url, token) {
  console.log('[Background] PSS Fetch URL:', url);

  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    }
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    console.error('[Background] PSS API error response:', errorText);
    throw new Error(`PSS API error: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function mapPssPriority(pssPriority) {
  if (pssPriority === 1 || pssPriority === 'Urgent') return { value: 1, label: 'Urgent' };
  if (pssPriority === 2 || pssPriority === 'High') return { value: 3, label: 'High' };
  if (pssPriority === 3 || pssPriority === 'Medium') return { value: 5, label: 'Medium' };
  if (pssPriority === 4 || pssPriority === 'Low') return { value: 9, label: 'Low' };
  return { value: 5, label: 'Medium' };
}

function sendProgressToTab(tabId, progress) {
  // Update extraction state with progress
  updateExtractionState({ progress });

  if (!tabId) return;
  chrome.tabs.sendMessage(tabId, {
    action: 'extractionProgress',
    progress
  }).catch(() => {}); // Ignore if tab closed
}

// ============================================
// GRAPH API - BASIC PLAN DATA
// ============================================

async function fetchBasicPlanData(planId, token, tabId) {
  console.log('[Background] Fetching basic plan data via Graph API...');

  sendProgressToTab(tabId, { status: 'fetching', message: 'Fetching plan via Graph API...' });

  // Get plan details
  const planResponse = await graphFetch(`/planner/plans/${planId}`, token);
  if (!planResponse.ok) {
    throw new Error(`Failed to fetch plan: ${planResponse.status}`);
  }
  const plan = await planResponse.json();

  // Get all tasks (with pagination)
  let tasks = [];
  let nextLink = `/planner/plans/${planId}/tasks`;
  while (nextLink) {
    const tasksResponse = await graphFetch(nextLink, token);
    if (!tasksResponse.ok) break;
    const tasksData = await tasksResponse.json();
    tasks = tasks.concat(tasksData.value || []);
    nextLink = tasksData['@odata.nextLink'];
  }

  sendProgressToTab(tabId, {
    status: 'extracting',
    message: `Found ${tasks.length} tasks, fetching details...`,
    total: tasks.length,
    current: 0
  });

  // Get buckets
  const bucketsResponse = await graphFetch(`/planner/plans/${planId}/buckets`, token);
  const bucketsData = await bucketsResponse.json();
  const buckets = bucketsData.value || [];

  const bucketMap = {};
  for (const bucket of buckets) {
    bucketMap[bucket.id] = bucket.name;
  }

  // Get task details (includes checklist, description)
  const detailsMap = {};
  for (let i = 0; i < tasks.length; i++) {
    const task = tasks[i];
    try {
      const detailResponse = await graphFetch(`/planner/tasks/${task.id}/details`, token);
      if (detailResponse.ok) {
        detailsMap[task.id] = await detailResponse.json();
      }
    } catch (e) {
      // Skip details that fail
    }

    if (i % 5 === 0) {
      sendProgressToTab(tabId, {
        status: 'extracting',
        message: `Fetching task details ${i + 1}/${tasks.length}...`,
        total: tasks.length,
        current: i
      });
    }
  }

  sendProgressToTab(tabId, {
    status: 'complete',
    message: `Fetched ${tasks.length} tasks via Graph API`,
    total: tasks.length,
    current: tasks.length
  });

  return {
    plan,
    tasks,
    buckets,
    bucketMap,
    detailsMap,
    source: 'graph-api',
    planType: 'basic',
    extractionMethod: 'graph-api',
    extractedAt: new Date().toISOString()
  };
}

// ============================================
// PSS API - PREMIUM PLAN DATA
// ============================================

// Open a PSS project session - required before making any data calls
async function openProjectSession(dynamicsOrg, planId, token) {
  console.log('[Background] Opening project session...');
  console.log('[Background] Dynamics Org:', dynamicsOrg);
  console.log('[Background] Plan ID:', planId);

  const response = await fetch(`${PSS_API}/xrm/openproject`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      XrmUrl: `https://${dynamicsOrg}`,
      XrmProjectId: planId
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[Background] Failed to open project session:', errorText);
    throw new Error(`Failed to open project: ${response.status} - ${errorText}`);
  }

  // CRITICAL: Get the location header - this is the base URL for all subsequent calls
  const locationUrl = response.headers.get('location');
  console.log('[Background] Session opened. Location URL:', locationUrl);

  const body = await response.json();

  return {
    baseUrl: locationUrl,
    sessionCapabilities: body.sessionCapabilities,
    project: body.project
  };
}

async function fetchPremiumPlanData(projectId, token, tabId) {
  console.log('[Background] Fetching premium plan data via PSS API');
  console.log('[Background] Project ID:', projectId);
  console.log('[Background] Token (first 50 chars):', token?.substring(0, 50));

  if (!projectId) {
    throw new Error('No PSS project ID provided');
  }

  // Parse project ID: msxrm_{dynamicsOrg}_{planId}
  let dynamicsOrg, planId;
  if (projectId.startsWith('msxrm_')) {
    const parts = projectId.substring(6); // Remove 'msxrm_'
    const lastUnderscoreIdx = parts.lastIndexOf('_');
    if (lastUnderscoreIdx > 0) {
      dynamicsOrg = parts.substring(0, lastUnderscoreIdx);
      planId = parts.substring(lastUnderscoreIdx + 1);
    }
  }

  if (!dynamicsOrg || !planId) {
    throw new Error(`Invalid project ID format: ${projectId}`);
  }

  sendProgressToTab(tabId, { status: 'fetching', message: 'Opening project session...' });

  // Step 1: Open project session to get the location URL
  const session = await openProjectSession(dynamicsOrg, planId, token);
  const baseUrl = session.baseUrl;

  if (!baseUrl) {
    throw new Error('No location URL returned from openproject');
  }

  sendProgressToTab(tabId, { status: 'fetching', message: 'Fetching plan data...' });

  console.log('[Background] Using base URL:', baseUrl);

  // Step 2: Fetch all data using the session's location URL
  const authHeaders = {
    'Authorization': `Bearer ${token}`,
    'Accept': 'application/json'
  };

  const [
    tasksResult,
    bucketsResult,
    resourcesResult,
    assignmentsResult,
    checklistsResult,
    labelsResult
  ] = await Promise.allSettled([
    fetch(`${baseUrl}/tasks/`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status)),
    fetch(`${baseUrl}/buckets`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status)),
    fetch(`${baseUrl}/resources/`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status)),
    fetch(`${baseUrl}/assignments/`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status)),
    fetch(`${baseUrl}/checklistItems`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status)),
    fetch(`${baseUrl}/labels`, { headers: authHeaders }).then(r => r.ok ? r.json() : Promise.reject(r.status))
  ]);

  // Log results
  console.log('[Background] PSS API results status:', {
    tasks: tasksResult.status,
    buckets: bucketsResult.status,
    resources: resourcesResult.status,
    assignments: assignmentsResult.status,
    checklists: checklistsResult.status,
    labels: labelsResult.status
  });

  if (tasksResult.status === 'rejected') {
    console.error('[Background] Tasks fetch failed:', tasksResult.reason);
  }

  // Log the actual response structure to understand the format
  if (tasksResult.status === 'fulfilled') {
    console.log('[Background] Tasks response structure:', tasksResult.value);
    console.log('[Background] Tasks response keys:', Object.keys(tasksResult.value || {}));
  }

  // Extract successful results - try both .value (OData) and direct array
  const extractData = (result) => {
    if (result.status !== 'fulfilled') return [];
    const data = result.value;
    // OData format: { value: [...] }
    if (data && Array.isArray(data.value)) return data.value;
    // Direct array format
    if (Array.isArray(data)) return data;
    // If it's an object with items
    if (data && typeof data === 'object') {
      console.log('[Background] Response data:', data);
    }
    return [];
  };

  const tasks = extractData(tasksResult);
  const buckets = extractData(bucketsResult);
  const resources = extractData(resourcesResult);
  const assignments = extractData(assignmentsResult);
  const checklists = extractData(checklistsResult);
  const labels = extractData(labelsResult);

  console.log('[Background] PSS API results:', {
    tasks: tasks.length,
    buckets: buckets.length,
    resources: resources.length,
    assignments: assignments.length,
    checklists: checklists.length,
    labels: labels.length
  });

  sendProgressToTab(tabId, {
    status: 'extracting',
    message: `Processing ${tasks.length} tasks...`,
    total: tasks.length,
    current: 0
  });

  // Build lookup maps
  const bucketMap = {};
  buckets.forEach(b => {
    // PSS API uses 'id' and 'name' directly
    bucketMap[b.id] = b.name;
  });

  // Build full resource map with all details
  const resourceMap = {};
  const resourceDetailMap = {};
  resources.forEach(r => {
    resourceMap[r.id] = r.name;
    resourceDetailMap[r.id] = {
      id: r.id,
      name: r.name,
      email: r.userPrincipalName || null,
      jobTitle: r.jobTitle || null,
      type: r.type || null
    };
  });

  // Group assignments by task with full resource details
  const taskAssignments = {};
  assignments.forEach(a => {
    const taskId = a.taskId;
    const resourceId = a.resourceId;
    if (!taskAssignments[taskId]) {
      taskAssignments[taskId] = [];
    }
    const resourceDetail = resourceDetailMap[resourceId] || {};
    taskAssignments[taskId].push({
      resourceId: resourceId,
      resourceName: resourceDetail.name || 'Unknown',
      email: resourceDetail.email || null,
      jobTitle: resourceDetail.jobTitle || null,
      percentWorkComplete: a.percentWorkComplete || 0
    });
  });

  // Group checklists by task
  const taskChecklists = {};
  checklists.forEach(c => {
    const taskId = c.taskId;
    if (!taskChecklists[taskId]) {
      taskChecklists[taskId] = [];
    }
    taskChecklists[taskId].push({
      id: c.id,
      title: c.name,
      isChecked: c.checked || false,
      order: c.order || 0
    });
  });

  // Map priority values
  const mapPriorityLabel = (priority) => {
    if (priority === 1) return 'Urgent';
    if (priority === 3) return 'Important';
    if (priority === 5) return 'Medium';
    if (priority === 9) return 'Low';
    return 'Medium';
  };

  // Enrich tasks with related data
  const enrichedTasks = tasks.map((task) => {
    const taskId = task.id;
    const assignmentDetails = taskAssignments[taskId] || [];

    return {
      id: taskId,
      title: task.name || 'Untitled Task',
      description: task.notes || task.unformattednotes || '',
      bucketId: task.bucketId,
      bucketName: bucketMap[task.bucketId] || '',
      startDateTime: task.scheduledStart || task.actualStart || task.start || null,
      dueDateTime: task.scheduledFinish || task.finish || null,
      percentComplete: task.percentComplete || 0,
      priority: task.priority || 5,
      priorityLabel: mapPriorityLabel(task.priority),
      isComplete: (task.percentComplete || 0) >= 100,
      isSummaryTask: task.summary || false,

      // Assignment details with names AND emails
      assignedTo: assignmentDetails.map(a => a.resourceName),
      assignedToEmails: assignmentDetails.map(a => a.email).filter(Boolean),
      assignments: assignmentDetails, // Full assignment details

      checklist: (taskChecklists[taskId] || []).sort((a, b) => (a.order || 0) - (b.order || 0)),
      duration: task.scheduledDuration ? `${Math.round(task.scheduledDuration / 3600)} hours` : '',

      // Hierarchy fields
      outlineLevel: task.outlineLevel || 0,
      outlineNumber: task.outlineNumber || '',
      parentId: task.parentId || null,
      order: task.order || task.index || 0,

      isMilestone: task.milestone || false,
      isCritical: task.critical || false,
      source: 'pss-api',
      // Keep raw data for advanced use
      raw: task
    };
  });

  // Sort tasks by order to maintain hierarchy
  enrichedTasks.sort((a, b) => a.order - b.order);

  sendProgressToTab(tabId, {
    status: 'complete',
    message: `Fetched ${enrichedTasks.length} tasks via PSS API`,
    total: enrichedTasks.length,
    current: enrichedTasks.length
  });

  return {
    tasks: enrichedTasks,
    buckets: buckets.map(b => ({
      id: b.id,
      name: b.name,
      order: b.order
    })),
    resources: resources.map(r => ({
      id: r.id,
      name: r.name,
      userPrincipalName: r.userPrincipalName,
      jobTitle: r.jobTitle
    })),
    labels: labels.map(l => ({
      id: l.id,
      name: l.name,
      color: l.color
    })),
    bucketMap,
    resourceMap,
    source: 'pss-api',
    planType: 'premium',
    extractionMethod: 'pss-api',
    extractedAt: new Date().toISOString()
  };
}

// ============================================
// TO DO API - MICROSOFT TO DO DATA
// ============================================

// Map To Do status to percentComplete
function mapToDoStatus(status) {
  const statusMap = {
    'notStarted': 0,
    'inProgress': 50,
    'completed': 100,
    'waitingOnOthers': 25,
    'deferred': 0
  };
  return statusMap[status] || 0;
}

// Map To Do importance to Planner-style priority
function mapToDoImportance(importance) {
  const importanceMap = {
    'low': { value: 9, label: 'Low' },
    'normal': { value: 5, label: 'Medium' },
    'high': { value: 3, label: 'High' }
  };
  return importanceMap[importance] || { value: 5, label: 'Medium' };
}

async function fetchToDoListData(listId, token, tabId) {
  console.log('[Background] Fetching To Do data via Graph API...');
  console.log('[Background] List ID:', listId);

  sendProgressToTab(tabId, { status: 'fetching', message: 'Fetching To Do lists...' });

  // If no listId provided, fetch all lists and their tasks
  let lists = [];
  let targetList = null;

  // Fetch all lists first
  const listsResponse = await graphFetch('/me/todo/lists', token);
  if (!listsResponse.ok) {
    throw new Error(`Failed to fetch To Do lists: ${listsResponse.status}`);
  }
  const listsData = await listsResponse.json();
  lists = listsData.value || [];

  console.log('[Background] Found', lists.length, 'To Do lists');

  // Find the target list or use all lists
  if (listId) {
    targetList = lists.find(l => l.id === listId);
    if (!targetList) {
      // Try to find by display name
      targetList = lists.find(l => l.displayName?.toLowerCase() === listId.toLowerCase());
    }
  }

  // If still no target list, use the first non-system list or all lists
  const listsToFetch = targetList ? [targetList] : lists;

  sendProgressToTab(tabId, {
    status: 'extracting',
    message: `Fetching tasks from ${listsToFetch.length} list(s)...`,
    total: listsToFetch.length,
    current: 0
  });

  let allTasks = [];
  const listMap = {};

  for (let i = 0; i < listsToFetch.length; i++) {
    const list = listsToFetch[i];
    listMap[list.id] = list.displayName;

    sendProgressToTab(tabId, {
      status: 'extracting',
      message: `Fetching "${list.displayName}"...`,
      total: listsToFetch.length,
      current: i
    });

    try {
      // Fetch tasks for this list (with checklist items expanded)
      let nextLink = `/me/todo/lists/${list.id}/tasks?$expand=checklistItems`;
      let listTasks = [];

      while (nextLink) {
        const tasksResponse = await graphFetch(nextLink, token);
        if (!tasksResponse.ok) {
          console.warn(`Failed to fetch tasks for list ${list.displayName}`);
          break;
        }
        const tasksData = await tasksResponse.json();
        listTasks = listTasks.concat(tasksData.value || []);
        nextLink = tasksData['@odata.nextLink'];
      }

      // Enrich tasks with list info
      listTasks.forEach(task => {
        task._listId = list.id;
        task._listName = list.displayName;
      });

      allTasks = allTasks.concat(listTasks);
    } catch (err) {
      console.error(`Error fetching tasks for list ${list.displayName}:`, err);
    }
  }

  sendProgressToTab(tabId, {
    status: 'processing',
    message: `Processing ${allTasks.length} tasks...`,
    total: allTasks.length,
    current: 0
  });

  // Transform tasks to common format
  const enrichedTasks = allTasks.map((task) => {
    const priority = mapToDoImportance(task.importance);
    const percentComplete = mapToDoStatus(task.status);

    return {
      id: task.id,
      title: task.title || 'Untitled Task',
      description: task.body?.content || '',
      bucketId: task._listId,
      bucketName: task._listName, // Use list name as bucket
      startDateTime: task.startDateTime?.dateTime || null,
      dueDateTime: task.dueDateTime?.dateTime || null,
      completedDateTime: task.completedDateTime?.dateTime || null,
      percentComplete: percentComplete,
      priority: priority.value,
      priorityLabel: priority.label,
      isComplete: task.status === 'completed',
      status: task.status,
      importance: task.importance,

      // To Do has no assignments
      assignedTo: [],
      assignments: [],

      // Checklist items
      checklist: (task.checklistItems || []).map(item => ({
        id: item.id,
        title: item.displayName,
        isChecked: item.isChecked || false
      })),

      // To Do specific fields
      isReminderOn: task.isReminderOn || false,
      reminderDateTime: task.reminderDateTime?.dateTime || null,
      hasAttachments: task.hasAttachments || false,
      categories: task.categories || [],

      // Meta
      createdDateTime: task.createdDateTime,
      lastModifiedDateTime: task.lastModifiedDateTime,
      source: 'todo-api',

      // Keep raw data
      raw: task
    };
  });

  sendProgressToTab(tabId, {
    status: 'complete',
    message: `Fetched ${enrichedTasks.length} tasks from To Do`,
    total: enrichedTasks.length,
    current: enrichedTasks.length
  });

  return {
    tasks: enrichedTasks,
    plan: targetList ? {
      id: targetList.id,
      title: targetList.displayName
    } : {
      id: 'all-lists',
      title: 'All To Do Lists'
    },
    // Use lists as "buckets" for compatibility
    buckets: listsToFetch.map(l => ({
      id: l.id,
      name: l.displayName,
      isShared: l.isShared || false,
      isOwner: l.isOwner !== false,
      wellknownListName: l.wellknownListName
    })),
    bucketMap: listMap,
    source: 'todo-api',
    serviceType: 'todo',
    planType: 'todo',
    extractionMethod: 'graph-api',
    extractedAt: new Date().toISOString()
  };
}

// ============================================
// MESSAGE HANDLERS
// ============================================

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  const tabId = sender.tab?.id;

  // Store token from content script
  if (request.action === 'storeToken') {
    const { type, token, metadata } = request;
    storeToken(type, token, metadata)
      .then(() => sendResponse({ success: true }))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }

  // Get all stored tokens
  if (request.action === 'getTokens') {
    getAllTokens()
      .then(tokens => sendResponse(tokens))
      .catch(err => sendResponse({ error: err.message }));
    return true;
  }

  // Get sender tab ID
  if (request.action === 'getTabId') {
    sendResponse({ tabId: tabId });
    return true;
  }

  // Get current extraction state
  if (request.action === 'getExtractionState') {
    sendResponse(extractionState);
    return true;
  }

  // DOM extraction started (from content.js)
  if (request.action === 'domExtractionStarted') {
    updateExtractionState({
      status: 'extracting',
      startedAt: Date.now(),
      error: null,
      method: request.method || 'dom',
      progress: { status: 'extracting', message: 'DOM extraction in progress...' }
    });
    sendResponse({ success: true });
    return true;
  }

  // DOM extraction completed (from content.js)
  if (request.action === 'domExtractionCompleted') {
    updateExtractionState({
      status: 'complete',
      completedAt: Date.now(),
      taskCount: request.taskCount || 0,
      progress: { status: 'complete', message: `Extracted ${request.taskCount || 0} tasks` }
    });
    sendResponse({ success: true });
    return true;
  }

  // DOM extraction failed (from content.js)
  if (request.action === 'domExtractionFailed') {
    updateExtractionState({
      status: 'error',
      error: request.error,
      completedAt: Date.now()
    });
    sendResponse({ success: true });
    return true;
  }

  // Progress update from content.js (DOM extraction)
  if (request.action === 'extractionProgress') {
    // Update state and badge
    updateExtractionState({ progress: request.progress });
    return false; // Don't send response, let other listeners handle
  }

  // Fetch basic plan data via Graph API
  if (request.action === 'fetchBasicPlan') {
    const { planId, token } = request;
    const targetTabId = request.tabId || tabId;

    // Update state to extracting
    updateExtractionState({
      status: 'extracting',
      startedAt: Date.now(),
      error: null,
      method: 'graph-api',
      progress: { status: 'fetching', message: 'Starting Graph API extraction...' }
    });

    fetchBasicPlanData(planId, token, targetTabId)
      .then(data => {
        updateExtractionState({
          status: 'complete',
          completedAt: Date.now(),
          taskCount: data.tasks?.length || 0,
          progress: { status: 'complete', message: `Extracted ${data.tasks?.length || 0} tasks` }
        });
        sendResponse({ success: true, data });
      })
      .catch(err => {
        console.error('[Background] fetchBasicPlan error:', err);
        updateExtractionState({
          status: 'error',
          error: err.message,
          completedAt: Date.now()
        });
        sendResponse({ success: false, error: err.message });
      });
    return true;
  }

  // Fetch To Do list data via Graph API
  if (request.action === 'fetchToDoList') {
    const { listId, token } = request;
    const targetTabId = request.tabId || tabId;

    // Update state to extracting
    updateExtractionState({
      status: 'extracting',
      startedAt: Date.now(),
      error: null,
      method: 'todo-api',
      progress: { status: 'fetching', message: 'Starting To Do extraction...' }
    });

    fetchToDoListData(listId, token, targetTabId)
      .then(data => {
        updateExtractionState({
          status: 'complete',
          completedAt: Date.now(),
          taskCount: data.tasks?.length || 0,
          progress: { status: 'complete', message: `Extracted ${data.tasks?.length || 0} tasks from To Do` }
        });
        sendResponse({ success: true, data });
      })
      .catch(err => {
        console.error('[Background] fetchToDoList error:', err);
        updateExtractionState({
          status: 'error',
          error: err.message,
          completedAt: Date.now()
        });
        sendResponse({ success: false, error: err.message });
      });
    return true;
  }

  // Fetch premium plan data via PSS API
  if (request.action === 'fetchPremiumPlan') {
    const { projectId, token } = request;
    const targetTabId = request.tabId || tabId;

    // Update state to extracting
    updateExtractionState({
      status: 'extracting',
      startedAt: Date.now(),
      error: null,
      method: 'pss-api',
      progress: { status: 'fetching', message: 'Starting PSS API extraction...' }
    });

    fetchPremiumPlanData(projectId, token, targetTabId)
      .then(data => {
        updateExtractionState({
          status: 'complete',
          completedAt: Date.now(),
          taskCount: data.tasks?.length || 0,
          progress: { status: 'complete', message: `Extracted ${data.tasks?.length || 0} tasks` }
        });
        sendResponse({ success: true, data });
      })
      .catch(err => {
        console.error('[Background] fetchPremiumPlan error:', err);
        updateExtractionState({
          status: 'error',
          error: err.message,
          completedAt: Date.now()
        });
        sendResponse({ success: false, error: err.message });
      });
    return true;
  }

  // ============================================
  // IMPORT HANDLERS
  // ============================================

  // Get import session with existing buckets and resources
  if (request.action === 'getImportSession') {
    (async () => {
      try {
        // Get stored PSS token and project ID
        const pssTokenData = await getToken('PSS');
        const pssProjectData = await getToken('PSS_PROJECT');

        if (!pssTokenData?.token || !pssProjectData?.projectId) {
          sendResponse({
            success: false,
            error: 'No PSS session found. Please navigate to a Premium Plan first.'
          });
          return;
        }

        const token = pssTokenData.token;
        const projectId = pssProjectData.projectId;

        // Parse project ID
        let dynamicsOrg, planId;
        if (projectId.startsWith('msxrm_')) {
          const parts = projectId.substring(6);
          const lastUnderscoreIdx = parts.lastIndexOf('_');
          if (lastUnderscoreIdx > 0) {
            dynamicsOrg = parts.substring(0, lastUnderscoreIdx);
            planId = parts.substring(lastUnderscoreIdx + 1);
          }
        }

        if (!dynamicsOrg || !planId) {
          sendResponse({ success: false, error: 'Invalid project ID format' });
          return;
        }

        // Open project session
        const session = await openProjectSession(dynamicsOrg, planId, token);

        // Fetch existing buckets and resources
        const authHeaders = {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        };

        const [bucketsResp, resourcesResp] = await Promise.allSettled([
          fetch(`${session.baseUrl}/buckets`, { headers: authHeaders }).then(r => r.ok ? r.json() : []),
          fetch(`${session.baseUrl}/resources/`, { headers: authHeaders }).then(r => r.ok ? r.json() : [])
        ]);

        const buckets = bucketsResp.status === 'fulfilled' ?
          (Array.isArray(bucketsResp.value) ? bucketsResp.value : bucketsResp.value?.value || []) : [];
        const resources = resourcesResp.status === 'fulfilled' ?
          (Array.isArray(resourcesResp.value) ? resourcesResp.value : resourcesResp.value?.value || []) : [];

        sendResponse({
          success: true,
          data: {
            baseUrl: session.baseUrl,
            token: token,
            buckets: buckets.map(b => ({ id: b.id, name: b.name })),
            resources: resources.map(r => ({
              id: r.id,
              name: r.name,
              userPrincipalName: r.userPrincipalName
            })),
            plannerUrl: `https://tasks.office.com` // Generic URL
          }
        });
      } catch (error) {
        console.error('[Background] getImportSession error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Create a task for import
  if (request.action === 'createImportTask') {
    const { taskData, baseUrl, token } = request;

    (async () => {
      try {
        const response = await fetch(`${baseUrl}/tasks/`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify(taskData)
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to create task: ${response.status} - ${errorText}`);
        }

        const createdTask = await response.json();
        sendResponse({ success: true, data: createdTask });
      } catch (error) {
        console.error('[Background] createImportTask error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Set task parent (for hierarchy)
  if (request.action === 'setTaskParent') {
    const { taskId, parentId, baseUrl, token } = request;

    (async () => {
      try {
        const response = await fetch(`${baseUrl}/tasks(${taskId})`, {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({ parentId: parentId })
        });

        if (!response.ok && response.status !== 204) {
          const errorText = await response.text();
          throw new Error(`Failed to set parent: ${response.status} - ${errorText}`);
        }

        sendResponse({ success: true });
      } catch (error) {
        console.error('[Background] setTaskParent error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Create checklist item for import
  if (request.action === 'createImportChecklist') {
    const { taskId, name, baseUrl, token } = request;

    (async () => {
      try {
        const response = await fetch(`${baseUrl}/tasks(${taskId})/checklistItems`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            name: name,
            completed: false
          })
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to create checklist: ${response.status} - ${errorText}`);
        }

        const createdItem = await response.json();
        sendResponse({ success: true, data: createdItem });
      } catch (error) {
        console.error('[Background] createImportChecklist error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Open results page
  if (request.action === 'openResults') {
    chrome.tabs.create({ url: chrome.runtime.getURL('results.html') });
    sendResponse({ success: true });
    return true;
  }

  // Download file
  if (request.action === 'downloadFile') {
    const { content, filename, mimeType } = request;
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);

    chrome.downloads.download({
      url: url,
      filename: filename,
      saveAs: true
    }, (downloadId) => {
      URL.revokeObjectURL(url);
      sendResponse({ success: true, downloadId });
    });

    return true;
  }
});

// ============================================
// LIFECYCLE
// ============================================

chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('[Planner Exporter] Extension installed');
  } else if (details.reason === 'update') {
    console.log('[Planner Exporter] Extension updated to version', chrome.runtime.getManifest().version);
  }
});

// Keep service worker alive
chrome.alarms.create('keepAlive', { periodInMinutes: 1 });
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'keepAlive') {
    // Periodic keep-alive ping
  }
});

console.log('[Planner Exporter] Background service worker initialized');
