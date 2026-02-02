/**
 * Background Service Worker for Planner Exporter
 * Handles API calls, token storage, and file downloads
 */

// ============================================
// API CONFIGURATION
// ============================================

const GRAPH_API = 'https://graph.microsoft.com/v1.0';
const PSS_API = 'https://project.microsoft.com/pss/api/v1.0';
const TODO_SUBSTRATE_API = 'https://substrate.office.com/todob2/api/v1';

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
  TODO: 'todoSubstrateToken',
  TODO_TIMESTAMP: 'todoSubstrateTokenTimestamp',
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

// Format a date string to ISO 8601 for Graph API
function formatDateForGraph(dateStr) {
  if (!dateStr) return null;
  try {
    // Handle various formats: YYYY-MM-DD, MM/DD/YYYY, etc.
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return null;
    return d.toISOString();
  } catch (e) {
    return null;
  }
}

// Generate a UUID v4 for Graph API checklist item keys
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
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
// TO DO API - MICROSOFT TO DO DATA (Substrate API)
// ============================================

// Extract user email from JWT token (for X-AnchorMailbox fallback)
function extractEmailFromToken(token) {
  try {
    if (!token || typeof token !== 'string') return null;
    const parts = token.split('.');
    if (parts.length !== 3) return null;

    // Decode the payload (second part)
    const payload = JSON.parse(atob(parts[1]));

    // Try common claims that contain the user's email
    const email = payload.upn ||
                  payload.unique_name ||
                  payload.preferred_username ||
                  payload.email ||
                  payload.sub; // sub sometimes contains email for user tokens

    // Validate it looks like an email
    if (email && email.includes('@')) {
      console.log('[Background] Extracted email from token:', email);
      return email;
    }
    return null;
  } catch (e) {
    console.log('[Background] Could not extract email from token:', e.message);
    return null;
  }
}

// Map To Do status to percentComplete
function mapToDoStatus(status) {
  // Substrate API uses different status values
  const statusMap = {
    'notStarted': 0,
    'NotStarted': 0,
    'inProgress': 50,
    'InProgress': 50,
    'completed': 100,
    'Completed': 100,
    'waitingOnOthers': 25,
    'WaitingOnOthers': 25,
    'deferred': 0,
    'Deferred': 0
  };
  return statusMap[status] || 0;
}

// Map To Do importance to Planner-style priority
function mapToDoImportance(importance) {
  // Substrate API may use different casing
  const normalizedImportance = (importance || '').toLowerCase();
  const importanceMap = {
    'low': { value: 9, label: 'Low' },
    'normal': { value: 5, label: 'Medium' },
    'high': { value: 3, label: 'High' }
  };
  return importanceMap[normalizedImportance] || { value: 5, label: 'Medium' };
}

// Get fresh To Do token and headers from storage or active tab
async function getFreshToDoToken() {
  // First try chrome.storage.local
  const data = await chrome.storage.local.get(['todoSubstrateToken', 'todoSubstrateTokenTimestamp', 'todoAnchorMailbox']);
  if (data.todoSubstrateToken) {
    const tokenAge = Date.now() - (data.todoSubstrateTokenTimestamp || 0);
    // Use 45 minutes threshold - tokens typically expire after ~1 hour
    const MAX_TOKEN_AGE = 45 * 60 * 1000; // 45 minutes
    if (tokenAge < MAX_TOKEN_AGE) {
      console.log('[Background] Using stored To Do token (age:', Math.round(tokenAge / 1000), 'seconds)');
      console.log('[Background] X-AnchorMailbox:', data.todoAnchorMailbox || 'not set');
      return {
        token: data.todoSubstrateToken,
        anchorMailbox: data.todoAnchorMailbox || null
      };
    }
    console.log('[Background] Stored token is too old:', Math.round(tokenAge / 1000), 'seconds, max age:', MAX_TOKEN_AGE / 1000);
  }

  // Fallback: Try to get token from active To Do tab
  console.log('[Background] No valid token in storage, trying active To Do tab...');
  try {
    const tabs = await chrome.tabs.query({ url: '*://to-do.office.com/*' });
    for (const tab of tabs) {
      try {
        const response = await chrome.tabs.sendMessage(tab.id, { action: 'getContext' });
        if (response && response.token) {
          console.log('[Background] Got token from active To Do tab');
          // Store it for future use
          await chrome.storage.local.set({
            todoSubstrateToken: response.token,
            todoSubstrateTokenTimestamp: Date.now()
          });
          return {
            token: response.token,
            anchorMailbox: null
          };
        }
      } catch (e) {
        // Tab might not have content script loaded
      }
    }
  } catch (e) {
    console.log('[Background] Could not query To Do tabs:', e.message);
  }

  return null;
}

// Substrate API fetch wrapper with retry logic
// tokenData can be a string (token only) or object { token, anchorMailbox }
async function substrateFetch(endpoint, tokenData, maxRetries = 3) {
  const url = endpoint.startsWith('http') ? endpoint : `${TODO_SUBSTRATE_API}${endpoint}`;
  console.log('[Background] Substrate fetch:', url);

  // Handle both string token and token object
  let currentToken = typeof tokenData === 'string' ? tokenData : tokenData?.token;
  let anchorMailbox = typeof tokenData === 'object' ? tokenData?.anchorMailbox : null;
  let lastError = null;

  // If no anchorMailbox, try to extract email from token as fallback
  if (!anchorMailbox && currentToken) {
    anchorMailbox = extractEmailFromToken(currentToken);
    if (anchorMailbox) {
      console.log('[Background] Using email from token as X-AnchorMailbox fallback:', anchorMailbox);
    }
  }

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      // Build headers
      const headers = {
        'Authorization': `Bearer ${currentToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      };

      // Add X-AnchorMailbox if available (critical for Substrate API routing)
      if (anchorMailbox) {
        headers['X-AnchorMailbox'] = anchorMailbox;
        console.log('[Background] Using X-AnchorMailbox:', anchorMailbox);
      }

      const response = await fetch(url, { headers });

      if (response.status === 401) {
        console.log(`[Background] 401 Unauthorized, attempt ${attempt + 1}/${maxRetries}`);

        // Try to get fresh token and headers from storage
        const freshTokenData = await getFreshToDoToken();
        if (freshTokenData && freshTokenData.token && freshTokenData.token !== currentToken) {
          console.log('[Background] Got fresh token from storage, retrying...');
          currentToken = freshTokenData.token;
          if (freshTokenData.anchorMailbox) {
            anchorMailbox = freshTokenData.anchorMailbox;
          }
          // Wait a bit before retrying
          await new Promise(r => setTimeout(r, 500 * (attempt + 1)));
          continue;
        }

        // If no fresh token, wait longer and retry (token might be captured soon)
        if (attempt < maxRetries - 1) {
          console.log('[Background] Waiting for token capture...');
          await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
          const newerTokenData = await getFreshToDoToken();
          if (newerTokenData && newerTokenData.token) {
            currentToken = newerTokenData.token;
            if (newerTokenData.anchorMailbox) {
              anchorMailbox = newerTokenData.anchorMailbox;
            }
          }
          continue;
        }
      }

      if (!response.ok) {
        const errorText = await response.text().catch(() => '');
        console.error('[Background] Substrate API error:', response.status, errorText);
        throw new Error(`Substrate API error: ${response.status} ${response.statusText}`);
      }

      return response.json();
    } catch (err) {
      lastError = err;
      if (attempt < maxRetries - 1) {
        console.log(`[Background] Fetch error, retrying: ${err.message}`);
        await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
      }
    }
  }

  throw lastError || new Error('Max retries exceeded');
}

// tokenData can be string (token only) or object { token, anchorMailbox }
async function fetchToDoListData(listId, listName, tokenData, tabId) {
  console.log('[Background] Fetching To Do data via Substrate API...');
  console.log('[Background] List ID:', listId);
  console.log('[Background] List Name:', listName);

  sendProgressToTab(tabId, { status: 'fetching', message: 'Fetching To Do lists...' });

  // Fetch all task folders (lists) from Substrate API
  let lists = [];
  let targetList = null;

  try {
    const listsData = await substrateFetch('/taskfolders?maxPageSize=200', tokenData);
    // Substrate API uses PascalCase - 'Value' not 'value'
    lists = listsData.Value || listsData.value || listsData || [];
    console.log('[Background] Found', lists.length, 'To Do lists');
    console.log('[Background] Lists structure sample:', lists[0]);
  } catch (err) {
    console.error('[Background] Failed to fetch task folders:', err);
    throw new Error(`Failed to fetch To Do lists: ${err.message}`);
  }

  // Find the target list by name first (more reliable from DOM), then by ID
  if (listName) {
    // Try to find by name (Substrate uses 'Name' with capital N)
    targetList = lists.find(l =>
      (l.Name || l.DisplayName || l.name || l.displayName || '').toLowerCase() === listName.toLowerCase()
    );
    console.log('[Background] Found list by name:', targetList ? 'Yes' : 'No');
  }

  // If not found by name, try by ID
  if (!targetList && listId) {
    // Substrate uses PascalCase - 'Id' (capital I) for the folder ID
    targetList = lists.find(l => l.Id === listId || l.id === listId);
    console.log('[Background] Found list by ID:', targetList ? 'Yes' : 'No');
  }

  // If we have a list name but couldn't find it, throw an error
  if (listName && !targetList) {
    console.error('[Background] Could not find list:', listName);
    console.log('[Background] Available lists:', lists.map(l => l.Name || l.name));
    throw new Error(`List "${listName}" not found. Please navigate to a specific list.`);
  }

  // If no target list found and no name specified, use all lists (fallback)
  const listsToFetch = targetList ? [targetList] : lists;

  if (targetList) {
    const targetName = targetList.Name || targetList.name;
    console.log('[Background] Fetching only list:', targetName);
    sendProgressToTab(tabId, {
      status: 'extracting',
      message: `Fetching tasks from "${targetName}"...`,
      total: 1,
      current: 0
    });
  } else {
    console.log('[Background] No specific list - fetching all', lists.length, 'lists');
    sendProgressToTab(tabId, {
      status: 'extracting',
      message: `Fetching tasks from ${listsToFetch.length} list(s)...`,
      total: listsToFetch.length,
      current: 0
    });
  }

  let allTasks = [];
  const listMap = {};

  for (let i = 0; i < listsToFetch.length; i++) {
    const list = listsToFetch[i];
    // Substrate API uses PascalCase - prioritize those
    const folderId = list.Id || list.id;
    const folderName = list.Name || list.DisplayName || list.name || list.displayName || 'Unknown List';

    listMap[folderId] = folderName;

    sendProgressToTab(tabId, {
      status: 'extracting',
      message: `Fetching "${folderName}"...`,
      total: listsToFetch.length,
      current: i
    });

    try {
      // Fetch tasks for this folder from Substrate API
      const tasksData = await substrateFetch(`/taskfolders/${folderId}/tasks?maxPageSize=200`, tokenData);
      // Substrate API uses PascalCase - 'Value' not 'value'
      let listTasks = tasksData.Value || tasksData.value || tasksData || [];

      console.log(`[Background] Fetched ${listTasks.length} tasks from "${folderName}"`);
      if (listTasks.length > 0) {
        console.log('[Background] Task structure sample:', listTasks[0]);
      }

      // Enrich tasks with list info
      listTasks.forEach(task => {
        task._listId = folderId;
        task._listName = folderName;
      });

      allTasks = allTasks.concat(listTasks);
    } catch (err) {
      console.error(`Error fetching tasks for folder ${folderName}:`, err);
    }
  }

  sendProgressToTab(tabId, {
    status: 'processing',
    message: `Processing ${allTasks.length} tasks...`,
    total: allTasks.length,
    current: 0
  });

  // Transform tasks to common format
  // Substrate API uses PascalCase property names
  const enrichedTasks = allTasks.map((task) => {
    // Substrate API uses PascalCase - prioritize those
    const taskId = task.Id || task.id;
    const title = task.Subject || task.Name || task.Title || task.subject || task.name || task.title || 'Untitled Task';
    const body = task.Body || task.body || {};
    const description = body.Content || body.content || '';
    const status = task.Status || task.status || 'NotStarted';
    const importance = task.Importance || task.importance || 'Normal';

    // Dates - Substrate uses DueDate, StartDate (not DueDateTime)
    const startDate = task.StartDate || task.StartDateTime || task.startDate || task.startDateTime;
    const dueDate = task.DueDate || task.DueDateTime || task.dueDate || task.dueDateTime;
    const completedDate = task.CompletedDate || task.CompletedDateTime || task.completedDate || task.completedDateTime;

    const priority = mapToDoImportance(importance);
    const percentComplete = mapToDoStatus(status);

    // Checklist items - Substrate uses Subtasks (not Checklist)
    const checklistItems = task.Subtasks || task.SubTasks || task.Checklist || task.ChecklistItems ||
                           task.subtasks || task.checklist || task.checklistItems || [];

    return {
      id: taskId,
      title: title,
      description: description,
      bucketId: task._listId,
      bucketName: task._listName, // Use list name as bucket
      // Dates - Substrate may return DateTime as nested object { DateTime: "..." } or direct string
      startDateTime: (typeof startDate === 'object' ? (startDate?.DateTime || startDate?.dateTime) : startDate) || null,
      dueDateTime: (typeof dueDate === 'object' ? (dueDate?.DateTime || dueDate?.dateTime) : dueDate) || null,
      completedDateTime: (typeof completedDate === 'object' ? (completedDate?.DateTime || completedDate?.dateTime) : completedDate) || null,
      percentComplete: percentComplete,
      priority: priority.value,
      priorityLabel: priority.label,
      isComplete: status.toLowerCase() === 'completed',
      status: status,
      importance: importance,

      // To Do has no assignments
      assignedTo: [],
      assignments: [],

      // Checklist items (Subtasks) - Substrate uses Subject and IsCompleted
      checklist: (Array.isArray(checklistItems) ? checklistItems : []).map(item => ({
        id: item.Id || item.id,
        title: item.Subject || item.DisplayName || item.Name || item.Title || item.subject || item.displayName || item.name || item.title || '',
        isChecked: item.IsCompleted || item.IsChecked || item.Completed || item.Checked ||
                   item.isCompleted || item.isChecked || item.completed || item.checked || false
      })),

      // To Do specific fields - Substrate uses PascalCase
      isReminderOn: task.IsReminderOn || task.isReminderOn || false,
      reminderDateTime: task.ReminderDate || task.ReminderDateTime || task.reminderDateTime || null,
      hasAttachments: task.HasAttachments || task.hasAttachments || false,
      categories: task.Categories || task.categories || [],

      // Meta - Substrate uses PascalCase
      createdDateTime: task.CreatedDateTime || task.DateTimeCreated || task.createdDateTime,
      lastModifiedDateTime: task.LastModifiedDateTime || task.DateTimeLastModified || task.lastModifiedDateTime,
      source: 'todo-substrate-api',

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

  // Build return object with list info
  // Substrate API uses PascalCase
  const targetListId = targetList ? (targetList.Id || targetList.id) : 'all-lists';
  const targetListName = targetList
    ? (targetList.Name || targetList.DisplayName || targetList.name || targetList.displayName)
    : 'All To Do Lists';

  return {
    tasks: enrichedTasks,
    plan: {
      id: targetListId,
      title: targetListName
    },
    // Use lists as "buckets" for compatibility
    buckets: listsToFetch.map(l => ({
      id: l.Id || l.id,
      name: l.Name || l.DisplayName || l.name || l.displayName || 'Unknown',
      isShared: l.IsShared || l.isShared || false,
      isOwner: l.IsOwner !== false && l.isOwner !== false,
      wellknownListName: l.WellknownListName || l.wellknownListName
    })),
    bucketMap: listMap,
    source: 'todo-substrate-api',
    serviceType: 'todo',
    planType: 'todo',
    extractionMethod: 'substrate-api',
    extractedAt: new Date().toISOString()
  };
}

// ============================================
// TO DO API - CREATE TASKS (Substrate API)
// ============================================

// Map priority to To Do Importance
function mapPriorityToImportance(priority) {
  if (typeof priority === 'number') {
    if (priority <= 1) return 'High';
    if (priority >= 9) return 'Low';
    return 'Normal';
  }
  if (typeof priority === 'string') {
    const p = priority.toLowerCase();
    if (p === 'urgent' || p === 'high' || p === 'important') return 'High';
    if (p === 'low') return 'Low';
  }
  return 'Normal';
}

// Create a task in To Do
// tokenData can be string (token only) or object { token, anchorMailbox }
async function createToDoTask(tokenData, listId, taskData) {
  console.log('[Background] Creating To Do task in list:', listId);

  // Handle both string token and token object
  const token = typeof tokenData === 'string' ? tokenData : tokenData?.token;
  let anchorMailbox = typeof tokenData === 'object' ? tokenData?.anchorMailbox : null;

  // Fallback: extract email from token if no anchorMailbox
  if (!anchorMailbox && token) {
    anchorMailbox = extractEmailFromToken(token);
  }

  const payload = {
    Subject: taskData.title || taskData.Subject || 'Untitled Task',
    Importance: mapPriorityToImportance(taskData.priority || taskData.importance),
    Status: taskData.status || 'NotStarted'
  };

  // Add body/notes if provided
  if (taskData.notes || taskData.description || taskData.body) {
    payload.Body = {
      Content: taskData.notes || taskData.description || taskData.body,
      ContentType: 'Text'
    };
  }

  // Add due date if provided
  if (taskData.dueDate || taskData.dueDateTime) {
    const dueDate = taskData.dueDate || taskData.dueDateTime;
    payload.DueDate = {
      DateTime: typeof dueDate === 'string' ? dueDate : new Date(dueDate).toISOString(),
      TimeZone: 'UTC'
    };
  }

  // Add start date if provided
  if (taskData.startDate || taskData.startDateTime) {
    const startDate = taskData.startDate || taskData.startDateTime;
    payload.StartDate = {
      DateTime: typeof startDate === 'string' ? startDate : new Date(startDate).toISOString(),
      TimeZone: 'UTC'
    };
  }

  // Build headers
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };

  // Add X-AnchorMailbox if available
  if (anchorMailbox) {
    headers['X-AnchorMailbox'] = anchorMailbox;
    console.log('[Background] Using X-AnchorMailbox for createToDoTask:', anchorMailbox);
  }

  const response = await fetch(
    `${TODO_SUBSTRATE_API}/taskfolders/${listId}/tasks`,
    {
      method: 'POST',
      headers: headers,
      body: JSON.stringify(payload)
    }
  );

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    console.error('[Background] Create task error:', response.status, errorText);
    throw new Error(`Failed to create task: ${response.status} ${response.statusText}`);
  }

  return await response.json();
}

// Add a checklist item (subtask) to a task
// tokenData can be string (token only) or object { token, anchorMailbox }
async function addToDoChecklistItem(tokenData, taskId, stepText, isCompleted = false) {
  console.log('[Background] Adding checklist item to task:', taskId);

  // Handle both string token and token object
  const token = typeof tokenData === 'string' ? tokenData : tokenData?.token;
  let anchorMailbox = typeof tokenData === 'object' ? tokenData?.anchorMailbox : null;

  // Fallback: extract email from token if no anchorMailbox
  if (!anchorMailbox && token) {
    anchorMailbox = extractEmailFromToken(token);
  }

  // Build headers
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };

  // Add X-AnchorMailbox if available
  if (anchorMailbox) {
    headers['X-AnchorMailbox'] = anchorMailbox;
  }

  const response = await fetch(
    `${TODO_SUBSTRATE_API}/tasks/${taskId}/subtasks`,
    {
      method: 'POST',
      headers: headers,
      body: JSON.stringify({
        Subject: stepText,
        IsCompleted: isCompleted
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    console.error('[Background] Add checklist item error:', response.status, errorText);
    throw new Error(`Failed to add checklist item: ${response.status} ${response.statusText}`);
  }

  return await response.json();
}

// Import multiple tasks to To Do
// tokenData can be string (token only) or object { token, anchorMailbox }
async function importTasksToToDo(tokenData, listId, tasks, tabId) {
  console.log('[Background] Importing', tasks.length, 'tasks to To Do list:', listId);

  const results = {
    success: [],
    failed: []
  };

  for (let i = 0; i < tasks.length; i++) {
    const taskData = tasks[i];

    sendProgressToTab(tabId, {
      status: 'importing',
      message: `Creating task ${i + 1} of ${tasks.length}...`,
      total: tasks.length,
      current: i
    });

    try {
      // Create the main task
      const createdTask = await createToDoTask(tokenData, listId, taskData);
      const taskId = createdTask.Id || createdTask.id;

      // Add checklist items if any
      if (taskData.checklist && Array.isArray(taskData.checklist)) {
        for (const item of taskData.checklist) {
          const itemText = item.title || item.Subject || item.text || item;
          const isChecked = item.isChecked || item.IsCompleted || item.checked || false;
          try {
            await addToDoChecklistItem(tokenData, taskId, itemText, isChecked);
          } catch (err) {
            console.warn('[Background] Failed to add checklist item:', err.message);
          }
        }
      }

      results.success.push({
        original: taskData,
        created: createdTask
      });
    } catch (err) {
      console.error('[Background] Failed to create task:', taskData.title, err.message);
      results.failed.push({
        task: taskData,
        error: err.message
      });
    }

    // Small delay to avoid rate limiting
    if (i < tasks.length - 1) {
      await new Promise(r => setTimeout(r, 200));
    }
  }

  sendProgressToTab(tabId, {
    status: 'complete',
    message: `Imported ${results.success.length} tasks (${results.failed.length} failed)`,
    total: tasks.length,
    current: tasks.length
  });

  return results;
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

  // Fetch To Do list data via Substrate API
  if (request.action === 'fetchToDoList') {
    const { listId, listName } = request;
    let { token } = request;
    const targetTabId = request.tabId || tabId;

    // Try to get fresh token from storage if not provided or seems stale
    (async () => {
      let tokenData = null;
      if (!token) {
        console.log('[Background] No token provided, checking storage...');
        tokenData = await getFreshToDoToken();
        token = tokenData?.token;
      } else {
        // Token was provided, but we still need anchorMailbox
        const storedData = await chrome.storage.local.get(['todoAnchorMailbox']);
        tokenData = { token, anchorMailbox: storedData.todoAnchorMailbox };
      }

      if (!token) {
        updateExtractionState({
          status: 'error',
          error: 'No token available. Please interact with To Do page first.',
          completedAt: Date.now()
        });
        sendResponse({ success: false, error: 'No token available. Please interact with the To Do page (scroll, click) to capture authentication.' });
        return;
      }

      // Update state to extracting
      updateExtractionState({
        status: 'extracting',
        startedAt: Date.now(),
        error: null,
        method: 'todo-substrate-api',
        progress: { status: 'fetching', message: `Fetching "${listName || 'To Do'}" list...` }
      });

      // Pass tokenData (with anchorMailbox) instead of just token
      fetchToDoListData(listId, listName, tokenData, targetTabId)
        .then(data => {
          const listTitle = data.plan?.title || listName || 'To Do';
          updateExtractionState({
            status: 'complete',
            completedAt: Date.now(),
            taskCount: data.tasks?.length || 0,
            progress: { status: 'complete', message: `Extracted ${data.tasks?.length || 0} tasks from "${listTitle}"` }
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
    })();
    return true;
  }

  // Create a single To Do task
  if (request.action === 'createToDoTask') {
    const { listId, taskData, anchorMailbox: requestAnchorMailbox } = request;
    let { token } = request;

    (async () => {
      // Get token and headers from storage if not provided
      let tokenData = null;
      if (!token) {
        tokenData = await getFreshToDoToken();
        token = tokenData?.token;
      } else {
        // Use anchorMailbox from request if provided, otherwise get from storage
        const storedData = await chrome.storage.local.get(['todoAnchorMailbox']);
        tokenData = { token, anchorMailbox: requestAnchorMailbox || storedData.todoAnchorMailbox };
      }
      if (!token) {
        sendResponse({ success: false, error: 'No token available' });
        return;
      }

      try {
        const createdTask = await createToDoTask(tokenData, listId, taskData);
        sendResponse({ success: true, data: createdTask });
      } catch (err) {
        console.error('[Background] createToDoTask error:', err);
        sendResponse({ success: false, error: err.message });
      }
    })();
    return true;
  }

  // Add checklist item to a To Do task
  if (request.action === 'addToDoChecklistItem') {
    const { taskId, text, isCompleted, anchorMailbox: requestAnchorMailbox } = request;
    let { token } = request;

    (async () => {
      let tokenData = null;
      if (!token) {
        tokenData = await getFreshToDoToken();
        token = tokenData?.token;
      } else {
        // Use anchorMailbox from request if provided, otherwise get from storage
        const storedData = await chrome.storage.local.get(['todoAnchorMailbox']);
        tokenData = { token, anchorMailbox: requestAnchorMailbox || storedData.todoAnchorMailbox };
      }
      if (!token) {
        sendResponse({ success: false, error: 'No token available' });
        return;
      }

      try {
        const item = await addToDoChecklistItem(tokenData, taskId, text, isCompleted || false);
        sendResponse({ success: true, data: item });
      } catch (err) {
        console.error('[Background] addToDoChecklistItem error:', err);
        sendResponse({ success: false, error: err.message });
      }
    })();
    return true;
  }

  // Import multiple tasks to To Do
  if (request.action === 'importTasksToToDo') {
    const { listId, tasks } = request;
    let { token } = request;
    const targetTabId = request.tabId || tabId;

    (async () => {
      let tokenData = null;
      if (!token) {
        tokenData = await getFreshToDoToken();
        token = tokenData?.token;
      } else {
        const storedData = await chrome.storage.local.get(['todoAnchorMailbox']);
        tokenData = { token, anchorMailbox: storedData.todoAnchorMailbox };
      }
      if (!token) {
        sendResponse({ success: false, error: 'No token available' });
        return;
      }

      try {
        const results = await importTasksToToDo(tokenData, listId, tasks, targetTabId);
        sendResponse({ success: true, data: results });
      } catch (err) {
        console.error('[Background] importTasksToToDo error:', err);
        sendResponse({ success: false, error: err.message });
      }
    })();
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

  // Get To Do import session (list of task lists)
  if (request.action === 'getToDoImportSession') {
    (async () => {
      try {
        // Get stored To Do token and headers
        const tokenData = await getFreshToDoToken();

        if (!tokenData || !tokenData.token) {
          // Check if there's an expired token to give better guidance
          const data = await chrome.storage.local.get(['todoSubstrateToken', 'todoSubstrateTokenTimestamp']);
          if (data.todoSubstrateToken) {
            const tokenAge = Math.round((Date.now() - (data.todoSubstrateTokenTimestamp || 0)) / 1000 / 60);
            sendResponse({
              success: false,
              error: `Token expired (${tokenAge} min old). Please open to-do.office.com, interact with the page (scroll, click a task), then try again.`
            });
          } else {
            sendResponse({
              success: false,
              error: 'No token found. Please open to-do.office.com and interact with the page first.'
            });
          }
          return;
        }

        // Fetch all task folders (lists) - pass full tokenData object for headers
        const listsData = await substrateFetch('/taskfolders?maxPageSize=200', tokenData);
        const lists = listsData.Value || listsData.value || [];

        console.log('[Background] getToDoImportSession: Found', lists.length, 'lists');

        sendResponse({
          success: true,
          data: {
            serviceType: 'todo',
            token: tokenData.token,
            anchorMailbox: tokenData.anchorMailbox,
            lists: lists.map(l => ({
              id: l.Id || l.id,
              name: l.Name || l.DisplayName || l.name || l.displayName || 'Unknown',
              isShared: l.IsShared || l.isShared || false,
              wellknownListName: l.WellknownListName || l.wellknownListName
            })),
            todoUrl: 'https://to-do.office.com'
          }
        });
      } catch (error) {
        console.error('[Background] getToDoImportSession error:', error);
        // Provide actionable error message
        if (error.message.includes('401') || error.message.includes('Max retries')) {
          sendResponse({
            success: false,
            error: 'Authentication failed. Please open to-do.office.com, scroll or click on tasks, then try again.'
          });
        } else {
          sendResponse({ success: false, error: error.message });
        }
      }
    })();
    return true;
  }

  // Get Planner Basic import session (uses Graph API)
  if (request.action === 'getBasicImportSession') {
    (async () => {
      try {
        // Get stored Graph token
        const graphTokenData = await getToken('GRAPH');

        if (!graphTokenData?.token) {
          sendResponse({
            success: false,
            error: 'No Graph API token found. Please navigate to a Basic Plan in Planner first.'
          });
          return;
        }

        // Get stored basic plan ID
        const storageData = await chrome.storage.local.get(['plannerBasicPlanId']);
        const planId = storageData.plannerBasicPlanId;

        if (!planId) {
          sendResponse({
            success: false,
            error: 'No Basic Plan detected. Please open a Basic Plan in Planner and interact with it first.'
          });
          return;
        }

        const token = graphTokenData.token;

        // Fetch buckets via Graph API
        const bucketsResponse = await graphFetch(`/planner/plans/${planId}/buckets`, token);
        if (!bucketsResponse.ok) {
          const errorText = await bucketsResponse.text().catch(() => '');
          throw new Error(`Failed to fetch buckets: ${bucketsResponse.status} ${errorText}`);
        }
        const bucketsData = await bucketsResponse.json();
        const buckets = bucketsData.value || [];

        sendResponse({
          success: true,
          data: {
            planId: planId,
            token: token,
            buckets: buckets.map(b => ({ id: b.id, name: b.name, orderHint: b.orderHint })),
            plannerUrl: 'https://tasks.office.com'
          }
        });
      } catch (error) {
        console.error('[Background] getBasicImportSession error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Create a task in Planner Basic via Graph API
  if (request.action === 'createBasicImportTask') {
    const { planId, token, taskData } = request;

    (async () => {
      try {
        const payload = {
          planId: planId,
          title: taskData.title || 'Untitled Task',
          priority: taskData.priority || 5
        };

        if (taskData.bucketId) {
          payload.bucketId = taskData.bucketId;
        }

        if (taskData.dueDateTime) {
          payload.dueDateTime = formatDateForGraph(taskData.dueDateTime);
        }

        if (taskData.startDateTime) {
          payload.startDateTime = formatDateForGraph(taskData.startDateTime);
        }

        // Assignments - Graph API format: { "userId": { "@odata.type": "...", "orderHint": " !" } }
        // Note: assignedTo contains emails, but Graph API needs user IDs
        // For now, we skip assignments since resolving emails to IDs requires additional API calls
        // TODO: Add user lookup if needed

        const response = await graphFetch(`/planner/tasks`, token, {
          method: 'POST',
          body: JSON.stringify(payload)
        });

        if (!response.ok) {
          const errorText = await response.text().catch(() => '');
          throw new Error(`Failed to create task: ${response.status} ${errorText}`);
        }

        const createdTask = await response.json();
        sendResponse({ success: true, data: createdTask });
      } catch (error) {
        console.error('[Background] createBasicImportTask error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

  // Update task details (description + checklist) for Planner Basic
  if (request.action === 'updateBasicTaskDetails') {
    const { taskId, token, details } = request;

    (async () => {
      try {
        // First, GET the task details to obtain the @odata.etag
        const detailsResponse = await graphFetch(`/planner/tasks/${taskId}/details`, token);
        if (!detailsResponse.ok) {
          throw new Error(`Failed to get task details: ${detailsResponse.status}`);
        }
        const currentDetails = await detailsResponse.json();
        const etag = currentDetails['@odata.etag'];

        // Build the patch payload
        const patchPayload = {};

        if (details.description) {
          patchPayload.description = details.description;
        }

        if (details.checklistItems && details.checklistItems.length > 0) {
          patchPayload.checklist = {};
          for (const item of details.checklistItems) {
            const uuid = generateUUID();
            patchPayload.checklist[uuid] = {
              '@odata.type': 'microsoft.graph.plannerChecklistItem',
              title: item,
              isChecked: false
            };
          }
        }

        // PATCH the task details with If-Match header
        const patchResponse = await graphFetch(`/planner/tasks/${taskId}/details`, token, {
          method: 'PATCH',
          headers: {
            'If-Match': etag
          },
          body: JSON.stringify(patchPayload)
        });

        if (!patchResponse.ok && patchResponse.status !== 204) {
          const errorText = await patchResponse.text().catch(() => '');
          throw new Error(`Failed to update task details: ${patchResponse.status} ${errorText}`);
        }

        sendResponse({ success: true });
      } catch (error) {
        console.error('[Background] updateBasicTaskDetails error:', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  }

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
