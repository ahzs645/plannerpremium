/**
 * To Do Page API for Planner Exporter
 * This script runs in the page context where CORS doesn't apply
 * It makes Substrate API calls using the captured token and returns results via postMessage
 */

(function() {
  'use strict';

  const TODO_SUBSTRATE_API = 'https://substrate.office.com/todob2/api/v1';

  // Listen for API requests from content script
  window.addEventListener('message', async (event) => {
    if (event.source !== window) return;

    const { type, requestId } = event.data || {};

    // Only handle specific request types (NOT responses)
    const validRequestTypes = [
      'TODO_API_GET_LISTS',
      'TODO_API_GET_TASKS',
      'TODO_API_CREATE_TASK',
      'TODO_API_ADD_SUBTASK'
    ];

    if (!type || !validRequestTypes.includes(type)) return;

    console.log('[ToDoPageApi] Received request:', type, requestId);

    try {
      let result;

      switch (type) {
        case 'TODO_API_GET_LISTS':
          result = await getLists();
          break;

        case 'TODO_API_GET_TASKS':
          result = await getTasksForList(event.data.listId);
          break;

        case 'TODO_API_CREATE_TASK':
          result = await createTask(event.data.listId, event.data.taskData);
          break;

        case 'TODO_API_ADD_SUBTASK':
          result = await addSubtask(event.data.taskId, event.data.text, event.data.isCompleted);
          break;

        default:
          throw new Error(`Unknown request type: ${type}`);
      }

      // Send success response
      window.postMessage({
        type: 'TODO_API_RESPONSE',
        requestId: requestId,
        success: true,
        data: result
      }, '*');

    } catch (error) {
      console.error('[ToDoPageApi] Error:', error);

      // Send error response
      window.postMessage({
        type: 'TODO_API_RESPONSE',
        requestId: requestId,
        success: false,
        error: error.message
      }, '*');
    }
  });

  // Get current token and headers
  function getAuthHeaders() {
    const token = window.__todoToken;
    const anchorMailbox = window.__todoAnchorMailbox;

    if (!token) {
      throw new Error('No token available. Please scroll or click on the page to capture authentication.');
    }

    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };

    if (anchorMailbox) {
      headers['X-AnchorMailbox'] = anchorMailbox;
    }

    return headers;
  }

  // Fetch all task lists
  async function getLists() {
    console.log('[ToDoPageApi] Fetching task lists...');

    const response = await fetch(`${TODO_SUBSTRATE_API}/taskfolders?maxPageSize=200`, {
      headers: getAuthHeaders()
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(`Failed to fetch lists: ${response.status} ${errorText}`);
    }

    const data = await response.json();
    const lists = data.Value || data.value || [];

    console.log('[ToDoPageApi] Found', lists.length, 'lists');

    return lists.map(l => ({
      id: l.Id || l.id,
      name: l.Name || l.DisplayName || l.name || l.displayName || 'Unknown',
      isShared: l.IsShared || l.isShared || false,
      wellknownListName: l.WellknownListName || l.wellknownListName
    }));
  }

  // Fetch tasks for a specific list
  async function getTasksForList(listId) {
    console.log('[ToDoPageApi] Fetching tasks for list:', listId);

    const response = await fetch(`${TODO_SUBSTRATE_API}/taskfolders/${listId}/tasks?maxPageSize=200`, {
      headers: getAuthHeaders()
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(`Failed to fetch tasks: ${response.status} ${errorText}`);
    }

    const data = await response.json();
    return data.Value || data.value || [];
  }

  // Create a task
  async function createTask(listId, taskData) {
    console.log('[ToDoPageApi] Creating task in list:', listId);

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

    const response = await fetch(`${TODO_SUBSTRATE_API}/taskfolders/${listId}/tasks`, {
      method: 'POST',
      headers: getAuthHeaders(),
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(`Failed to create task: ${response.status} ${errorText}`);
    }

    return await response.json();
  }

  // Add a subtask (checklist item)
  async function addSubtask(taskId, text, isCompleted = false) {
    console.log('[ToDoPageApi] Adding subtask to task:', taskId);

    const response = await fetch(`${TODO_SUBSTRATE_API}/tasks/${taskId}/subtasks`, {
      method: 'POST',
      headers: getAuthHeaders(),
      body: JSON.stringify({
        Subject: text,
        IsCompleted: isCompleted
      })
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(`Failed to add subtask: ${response.status} ${errorText}`);
    }

    return await response.json();
  }

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

  console.log('[ToDoPageApi] Page API initialized');
})();
