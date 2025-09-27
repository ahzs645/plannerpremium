/**
 * Background script for Microsoft Planner Interface
 * Handles data processing, storage, and API communication
 */

class PlannerBackgroundService {
  constructor() {
    this.plannerData = new Map();
    this.settings = {
      autoExtract: true,
      realTimeUpdates: false,
      dataRetentionDays: 30,
      maxStoredPlans: 50
    };

    this.init();
  }

  init() {
    this.setupEventListeners();
    this.loadSettings();
    this.cleanupOldData();
  }

  setupEventListeners() {
    // Listen for messages from content scripts and popup
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
      this.handleMessage(request, sender, sendResponse);
      return true; // Keep the message channel open for async responses
    });

    // Listen for tab updates to trigger auto-extraction
    chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
      this.handleTabUpdate(tabId, changeInfo, tab);
    });

    // Handle extension installation/startup
    chrome.runtime.onStartup.addListener(() => {
      this.loadSettings();
    });

    chrome.runtime.onInstalled.addListener((details) => {
      this.handleInstallation(details);
    });
  }

  async handleMessage(request, sender, sendResponse) {
    try {
      switch (request.action) {
        case 'plannerDataExtracted':
          await this.processPlannerData(request.data, sender.tab);
          sendResponse({ success: true });
          break;

        case 'getPlannerData':
          const data = await this.getPlannerData(request.planId);
          sendResponse({ success: true, data });
          break;

        case 'getAllPlans':
          const allPlans = await this.getAllPlans();
          sendResponse({ success: true, plans: allPlans });
          break;

        case 'exportData':
          const exportData = await this.exportData(request.format, request.planId);
          sendResponse({ success: true, data: exportData });
          break;

        case 'updateSettings':
          await this.updateSettings(request.settings);
          sendResponse({ success: true });
          break;

        case 'getSettings':
          sendResponse({ success: true, settings: this.settings });
          break;

        case 'clearData':
          await this.clearData(request.planId);
          sendResponse({ success: true });
          break;

        default:
          sendResponse({ success: false, error: 'Unknown action' });
      }
    } catch (error) {
      console.error('Error handling message:', error);
      sendResponse({ success: false, error: error.message });
    }
  }

  async processPlannerData(data, tab) {
    try {
      // Add metadata
      const processedData = {
        ...data,
        tabId: tab?.id,
        url: tab?.url,
        title: tab?.title,
        extractedAt: new Date().toISOString(),
        version: chrome.runtime.getManifest().version
      };

      // Generate plan ID from URL or plan name
      const planId = this.generatePlanId(processedData);
      processedData.planId = planId;

      // Store in memory
      this.plannerData.set(planId, processedData);

      // Store in chrome storage
      await this.storePlannerData(planId, processedData);

      // Process and analyze the data
      await this.analyzeData(processedData);

      // Send notification if significant changes detected
      await this.checkForChanges(planId, processedData);

      console.log(`Processed Planner data for plan: ${planId}`);
    } catch (error) {
      console.error('Error processing Planner data:', error);
    }
  }

  generatePlanId(data) {
    // Generate a unique ID based on plan name and URL
    const planName = data.planData?.planName || 'unknown-plan';
    const urlHash = this.hashCode(data.url || '');
    return `${planName.toLowerCase().replace(/[^a-z0-9]/g, '-')}-${urlHash}`;
  }

  hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32-bit integer
    }
    return Math.abs(hash).toString(36);
  }

  async storePlannerData(planId, data) {
    try {
      // Get existing data
      const result = await chrome.storage.local.get(['plannerPlans', 'plannerIndex']);
      const plans = result.plannerPlans || {};
      const index = result.plannerIndex || [];

      // Update plans
      plans[planId] = data;

      // Update index
      const existingIndex = index.findIndex(item => item.planId === planId);
      const indexEntry = {
        planId,
        planName: data.planData?.planName || 'Unknown Plan',
        url: data.url,
        lastUpdated: data.extractedAt,
        taskCount: data.taskData?.length || 0
      };

      if (existingIndex >= 0) {
        index[existingIndex] = indexEntry;
      } else {
        index.push(indexEntry);

        // Limit the number of stored plans
        if (index.length > this.settings.maxStoredPlans) {
          const oldestPlan = index.shift();
          delete plans[oldestPlan.planId];
        }
      }

      // Store updated data
      await chrome.storage.local.set({
        plannerPlans: plans,
        plannerIndex: index
      });

    } catch (error) {
      console.error('Error storing Planner data:', error);
    }
  }

  async analyzeData(data) {
    try {
      const analysis = {
        planId: data.planId,
        analyzedAt: new Date().toISOString(),
        metrics: {}
      };

      if (data.taskData && Array.isArray(data.taskData)) {
        const tasks = data.taskData;

        analysis.metrics = {
          totalTasks: tasks.length,
          completedTasks: tasks.filter(t => t.completed || t.progress === 100).length,
          inProgressTasks: tasks.filter(t => t.progress > 0 && t.progress < 100).length,
          notStartedTasks: tasks.filter(t => !t.progress || t.progress === 0).length,
          assignedTasks: tasks.filter(t => t.assignedTo).length,
          unassignedTasks: tasks.filter(t => !t.assignedTo).length,
          averageProgress: tasks.reduce((sum, t) => sum + (t.progress || 0), 0) / tasks.length,

          // Priority distribution
          priorities: this.groupBy(tasks, 'priority'),

          // Bucket distribution
          buckets: this.groupBy(tasks, 'bucket'),

          // Assignment distribution
          assignments: this.groupBy(tasks, 'assignedTo')
        };
      }

      // Store analysis
      await chrome.storage.local.set({
        [`analysis_${data.planId}`]: analysis
      });

      return analysis;
    } catch (error) {
      console.error('Error analyzing data:', error);
      return null;
    }
  }

  groupBy(array, key) {
    return array.reduce((groups, item) => {
      const value = item[key] || 'unspecified';
      groups[value] = (groups[value] || 0) + 1;
      return groups;
    }, {});
  }

  async checkForChanges(planId, newData) {
    try {
      const result = await chrome.storage.local.get([`previous_${planId}`]);
      const previousData = result[`previous_${planId}`];

      if (previousData) {
        const changes = this.detectChanges(previousData, newData);

        if (changes.hasSignificantChanges) {
          await this.sendChangeNotification(planId, changes);
        }
      }

      // Store current data as previous for next comparison
      await chrome.storage.local.set({
        [`previous_${planId}`]: newData
      });

    } catch (error) {
      console.error('Error checking for changes:', error);
    }
  }

  detectChanges(oldData, newData) {
    const changes = {
      hasSignificantChanges: false,
      taskChanges: {
        added: [],
        removed: [],
        modified: [],
        completed: []
      },
      planChanges: {}
    };

    // Compare tasks
    const oldTasks = oldData.taskData || [];
    const newTasks = newData.taskData || [];

    const oldTaskIds = new Set(oldTasks.map(t => t.id || t.name));
    const newTaskIds = new Set(newTasks.map(t => t.id || t.name));

    // Find added tasks
    newTasks.forEach(task => {
      const taskId = task.id || task.name;
      if (!oldTaskIds.has(taskId)) {
        changes.taskChanges.added.push(task);
        changes.hasSignificantChanges = true;
      }
    });

    // Find removed tasks
    oldTasks.forEach(task => {
      const taskId = task.id || task.name;
      if (!newTaskIds.has(taskId)) {
        changes.taskChanges.removed.push(task);
        changes.hasSignificantChanges = true;
      }
    });

    // Find modified/completed tasks
    newTasks.forEach(newTask => {
      const taskId = newTask.id || newTask.name;
      const oldTask = oldTasks.find(t => (t.id || t.name) === taskId);

      if (oldTask) {
        const progressChanged = (oldTask.progress || 0) !== (newTask.progress || 0);
        const statusChanged = oldTask.completed !== newTask.completed;

        if (progressChanged || statusChanged) {
          changes.taskChanges.modified.push({
            old: oldTask,
            new: newTask,
            changes: { progressChanged, statusChanged }
          });

          if (statusChanged && newTask.completed) {
            changes.taskChanges.completed.push(newTask);
          }

          changes.hasSignificantChanges = true;
        }
      }
    });

    return changes;
  }

  async sendChangeNotification(planId, changes) {
    // Create a notification about the changes
    const notificationOptions = {
      type: 'basic',
      iconUrl: 'icon48.png',
      title: 'Planner Update Detected',
      message: this.generateChangeMessage(changes)
    };

    const notificationsApi = chrome.notifications;
    if (!notificationsApi || typeof notificationsApi.create !== 'function') {
      return;
    }

    try {
      await notificationsApi.create(
        `planner_change_${planId}_${Date.now()}`,
        notificationOptions
      );
    } catch (error) {
      console.error('Error sending notification:', error);
    }
  }

  generateChangeMessage(changes) {
    const messages = [];

    if (changes.taskChanges.added.length > 0) {
      messages.push(`${changes.taskChanges.added.length} task(s) added`);
    }

    if (changes.taskChanges.completed.length > 0) {
      messages.push(`${changes.taskChanges.completed.length} task(s) completed`);
    }

    if (changes.taskChanges.modified.length > 0) {
      messages.push(`${changes.taskChanges.modified.length} task(s) updated`);
    }

    return messages.join(', ') || 'Plan updated';
  }

  async handleTabUpdate(tabId, changeInfo, tab) {
    if (!this.settings.autoExtract) return;

    // Check if the tab is a Planner page and has finished loading
    if (changeInfo.status === 'complete' && tab.url) {
      const isPlannerPage = tab.url.includes('planner.cloud.microsoft') ||
                           tab.url.includes('tasks.office.com');

      if (isPlannerPage) {
        // Small delay to ensure page is fully rendered
        setTimeout(async () => {
          try {
            await chrome.tabs.sendMessage(tabId, { action: 'extractPlannerData' });
          } catch (error) {
            // Content script might not be ready yet
            console.log('Content script not ready for auto-extraction');
          }
        }, 2000);
      }
    }
  }

  async handleInstallation(details) {
    if (details.reason === 'install') {
      // Set default settings
      await chrome.storage.local.set({
        settings: this.settings,
        plannerPlans: {},
        plannerIndex: []
      });

      // Create welcome notification
      const notificationsApi = chrome.notifications;
      if (notificationsApi && typeof notificationsApi.create === 'function') {
        try {
          await notificationsApi.create('planner_welcome', {
            type: 'basic',
            iconUrl: 'icon48.png',
            title: 'Planner Interface Installed',
            message: 'Visit a Microsoft Planner page to start extracting data!'
          });
        } catch (error) {
          console.error('Error showing welcome notification:', error);
        }
      }
    }
  }

  async loadSettings() {
    try {
      const result = await chrome.storage.local.get(['settings']);
      if (result.settings) {
        this.settings = { ...this.settings, ...result.settings };
      }
    } catch (error) {
      console.error('Error loading settings:', error);
    }
  }

  async updateSettings(newSettings) {
    try {
      this.settings = { ...this.settings, ...newSettings };
      await chrome.storage.local.set({ settings: this.settings });
    } catch (error) {
      console.error('Error updating settings:', error);
      throw error;
    }
  }

  async getPlannerData(planId) {
    try {
      if (this.plannerData.has(planId)) {
        return this.plannerData.get(planId);
      }

      const result = await chrome.storage.local.get(['plannerPlans']);
      const plans = result.plannerPlans || {};
      return plans[planId] || null;
    } catch (error) {
      console.error('Error getting Planner data:', error);
      return null;
    }
  }

  async getAllPlans() {
    try {
      const result = await chrome.storage.local.get(['plannerIndex']);
      return result.plannerIndex || [];
    } catch (error) {
      console.error('Error getting all plans:', error);
      return [];
    }
  }

  async exportData(format, planId) {
    try {
      let data;

      if (planId) {
        data = await this.getPlannerData(planId);
      } else {
        const result = await chrome.storage.local.get(['plannerPlans']);
        data = result.plannerPlans || {};
      }

      switch (format) {
        case 'json':
          return JSON.stringify(data, null, 2);

        case 'csv':
          return this.convertToCSV(data);

        default:
          throw new Error('Unsupported export format');
      }
    } catch (error) {
      console.error('Error exporting data:', error);
      throw error;
    }
  }

  convertToCSV(data) {
    // Implementation for CSV export
    const headers = ['Plan ID', 'Plan Name', 'Task Name', 'Assigned To', 'Progress', 'Status'];
    const rows = [];

    if (data.taskData && Array.isArray(data.taskData)) {
      data.taskData.forEach(task => {
        rows.push([
          data.planId || '',
          data.planData?.planName || '',
          task.name || '',
          task.assignedTo || '',
          task.progress || 0,
          task.completed ? 'Completed' : 'Active'
        ]);
      });
    }

    return [headers, ...rows]
      .map(row => row.map(field => `"${field}"`).join(','))
      .join('\n');
  }

  async clearData(planId) {
    try {
      if (planId) {
        // Clear specific plan
        this.plannerData.delete(planId);

        const result = await chrome.storage.local.get(['plannerPlans', 'plannerIndex']);
        const plans = result.plannerPlans || {};
        const index = result.plannerIndex || [];

        delete plans[planId];
        const newIndex = index.filter(item => item.planId !== planId);

        await chrome.storage.local.set({
          plannerPlans: plans,
          plannerIndex: newIndex
        });

        // Clear analysis and previous data
        await chrome.storage.local.remove([
          `analysis_${planId}`,
          `previous_${planId}`
        ]);
      } else {
        // Clear all data
        this.plannerData.clear();
        await chrome.storage.local.set({
          plannerPlans: {},
          plannerIndex: []
        });
      }
    } catch (error) {
      console.error('Error clearing data:', error);
      throw error;
    }
  }

  async cleanupOldData() {
    try {
      const result = await chrome.storage.local.get(['plannerIndex']);
      const index = result.plannerIndex || [];

      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - this.settings.dataRetentionDays);

      const plansToRemove = index.filter(plan =>
        new Date(plan.lastUpdated) < cutoffDate
      );

      for (const plan of plansToRemove) {
        await this.clearData(plan.planId);
      }

      if (plansToRemove.length > 0) {
        console.log(`Cleaned up ${plansToRemove.length} old plans`);
      }
    } catch (error) {
      console.error('Error cleaning up old data:', error);
    }
  }
}

// Initialize the background service
const plannerBackgroundService = new PlannerBackgroundService();

console.log('Microsoft Planner Background Service initialized');
