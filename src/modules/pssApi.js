/**
 * PSS (Project Scheduling Service) API Module
 * For Premium Plans (Project for the Web)
 * API Base: https://project.microsoft.com/pss/api/v1.0
 */

export const PSS_API_BASE = 'https://project.microsoft.com/pss/api/v1.0';

export const PssApi = {
  /**
   * Make an authenticated PSS API request
   */
  async fetch(endpoint, token, options = {}) {
    const url = endpoint.startsWith('http') ? endpoint : `${PSS_API_BASE}${endpoint}`;

    const response = await fetch(url, {
      ...options,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        ...options.headers
      }
    });

    if (!response.ok) {
      throw new Error(`PSS API error: ${response.status} ${response.statusText}`);
    }

    return response.json();
  },

  /**
   * Get all tasks for a project
   */
  async getTasks(projectId, token) {
    return this.fetch(`/projects('${projectId}')/tasks`, token);
  },

  /**
   * Get all buckets for a project
   */
  async getBuckets(projectId, token) {
    return this.fetch(`/projects('${projectId}')/buckets`, token);
  },

  /**
   * Get all resources (people) for a project
   */
  async getResources(projectId, token) {
    return this.fetch(`/projects('${projectId}')/resources`, token);
  },

  /**
   * Get task assignments
   */
  async getAssignments(projectId, token) {
    return this.fetch(`/projects('${projectId}')/assignments`, token);
  },

  /**
   * Get checklist items
   */
  async getChecklistItems(projectId, token) {
    return this.fetch(`/projects('${projectId}')/checklistItems`, token);
  },

  /**
   * Get labels
   */
  async getLabels(projectId, token) {
    return this.fetch(`/projects('${projectId}')/labels`, token);
  },

  /**
   * Get sprints
   */
  async getSprints(projectId, token) {
    return this.fetch(`/projects('${projectId}')/sprints`, token);
  },

  /**
   * Get task dependencies/links
   */
  async getLinks(projectId, token) {
    return this.fetch(`/projects('${projectId}')/links`, token);
  },

  /**
   * Get attachments
   */
  async getAttachments(projectId, token) {
    return this.fetch(`/projects('${projectId}')/attachments`, token);
  },

  /**
   * Get project details
   */
  async getProject(projectId, token) {
    return this.fetch(`/projects('${projectId}')`, token);
  },

  /**
   * Fetch all plan data in parallel for efficiency
   */
  async fetchAllPlanData(projectId, token, onProgress = null) {
    if (onProgress) {
      onProgress({ status: 'fetching', message: 'Fetching plan data via API...' });
    }

    // Fetch all data in parallel
    const [
      tasksResult,
      bucketsResult,
      resourcesResult,
      assignmentsResult,
      checklistsResult,
      labelsResult
    ] = await Promise.allSettled([
      this.getTasks(projectId, token),
      this.getBuckets(projectId, token),
      this.getResources(projectId, token),
      this.getAssignments(projectId, token),
      this.getChecklistItems(projectId, token),
      this.getLabels(projectId, token)
    ]);

    // Extract successful results
    const tasks = tasksResult.status === 'fulfilled' ? (tasksResult.value.value || []) : [];
    const buckets = bucketsResult.status === 'fulfilled' ? (bucketsResult.value.value || []) : [];
    const resources = resourcesResult.status === 'fulfilled' ? (resourcesResult.value.value || []) : [];
    const assignments = assignmentsResult.status === 'fulfilled' ? (assignmentsResult.value.value || []) : [];
    const checklists = checklistsResult.status === 'fulfilled' ? (checklistsResult.value.value || []) : [];
    const labels = labelsResult.status === 'fulfilled' ? (labelsResult.value.value || []) : [];

    // Build lookup maps
    const bucketMap = {};
    buckets.forEach(b => { bucketMap[b.msdyn_projectbucketid] = b.msdyn_name; });

    const resourceMap = {};
    resources.forEach(r => { resourceMap[r.msdyn_projectteamid] = r.msdyn_name; });

    // Group assignments by task
    const taskAssignments = {};
    assignments.forEach(a => {
      if (!taskAssignments[a.msdyn_projecttaskid]) {
        taskAssignments[a.msdyn_projecttaskid] = [];
      }
      taskAssignments[a.msdyn_projecttaskid].push({
        resourceId: a.msdyn_resourceid,
        resourceName: resourceMap[a.msdyn_resourceid] || 'Unknown'
      });
    });

    // Group checklists by task
    const taskChecklists = {};
    checklists.forEach(c => {
      if (!taskChecklists[c.msdyn_projecttaskid]) {
        taskChecklists[c.msdyn_projecttaskid] = [];
      }
      taskChecklists[c.msdyn_projecttaskid].push({
        id: c.msdyn_projecttaskchecklistitemid,
        title: c.msdyn_name,
        isChecked: c.msdyn_ischecked || false,
        order: c.msdyn_order
      });
    });

    // Enrich tasks with related data
    const enrichedTasks = tasks.map((task, index) => ({
      id: task.msdyn_projecttaskid,
      title: task.msdyn_subject || task.msdyn_name || 'Untitled Task',
      description: task.msdyn_description || '',
      bucketId: task.msdyn_projectbucketid,
      bucketName: bucketMap[task.msdyn_projectbucketid] || '',
      startDateTime: task.msdyn_scheduledstart || task.msdyn_actualstart || null,
      dueDateTime: task.msdyn_scheduledend || task.msdyn_duedate || null,
      percentComplete: task.msdyn_progress || 0,
      priority: mapPssPriority(task.msdyn_priority),
      priorityLabel: mapPssPriorityLabel(task.msdyn_priority),
      isComplete: (task.msdyn_progress || 0) >= 100,
      isSummaryTask: task.msdyn_issummary || false,
      assignedTo: (taskAssignments[task.msdyn_projecttaskid] || []).map(a => a.resourceName),
      checklist: (taskChecklists[task.msdyn_projecttaskid] || []).sort((a, b) => a.order - b.order),
      duration: task.msdyn_duration ? `${task.msdyn_duration} hours` : '',
      outlineLevel: task.msdyn_outlinelevel || 0,
      wbsId: task.msdyn_wbsid || '',
      source: 'pss-api'
    }));

    if (onProgress) {
      onProgress({
        status: 'complete',
        message: `Fetched ${enrichedTasks.length} tasks via API`,
        total: enrichedTasks.length,
        current: enrichedTasks.length
      });
    }

    return {
      tasks: enrichedTasks,
      buckets: buckets.map(b => ({
        id: b.msdyn_projectbucketid,
        name: b.msdyn_name
      })),
      resources: resources.map(r => ({
        id: r.msdyn_projectteamid,
        name: r.msdyn_name
      })),
      labels: labels.map(l => ({
        id: l.msdyn_projectlabelid,
        name: l.msdyn_name,
        color: l.msdyn_color
      })),
      bucketMap,
      resourceMap,
      extractionMethod: 'pss-api',
      extractedAt: new Date().toISOString()
    };
  }
};

// Priority mapping (PSS uses different values than Planner)
function mapPssPriority(pssPriority) {
  // PSS priority values may vary, adjust as needed
  if (pssPriority === 1 || pssPriority === 'Urgent') return 1;
  if (pssPriority === 2 || pssPriority === 'High') return 3;
  if (pssPriority === 3 || pssPriority === 'Medium') return 5;
  if (pssPriority === 4 || pssPriority === 'Low') return 9;
  return 5; // Default to medium
}

function mapPssPriorityLabel(pssPriority) {
  if (pssPriority === 1) return 'Urgent';
  if (pssPriority === 2) return 'High';
  if (pssPriority === 3) return 'Medium';
  if (pssPriority === 4) return 'Low';
  return 'Medium';
}

export default PssApi;
