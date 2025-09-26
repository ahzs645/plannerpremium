/**
 * Microsoft Planner Content Script
 * Extracts key information from Planner interface
 */

class PlannerDataExtractor {
  constructor() {
    this.planData = {};
    this.taskData = [];
    this.observers = [];
    this.updateTimeout = null;
  }

  async init() {
    if (document.readyState === 'loading') {
      await new Promise(resolve => {
        document.addEventListener('DOMContentLoaded', resolve);
      });
    }

    this.extractData();
    this.setupMutationObserver();
  }

  extractData() {
    this.extractPlanInfo();
    this.extractTaskInfo();
    this.extractBucketInfo();
    this.sendDataToBackground();
  }

  extractPlanInfo() {
    // Extract plan name from multiple possible locations
    this.extractPlanName();
    this.extractCurrentView();
    this.extractAccessLevel();
  }

  extractPlanName() {
    const titleCandidate = this.cleanPlanTitle(document.title);
    if (titleCandidate) {
      this.planData.planName = titleCandidate;
      return;
    }

    const planNameSelectors = [
      'header [role="heading"]',
      'header h1',
      'header h2',
      '[role="heading"][aria-level="1"]',
      '[data-testid*="plan"]',
      '[aria-label*="Plan name" i]',
      '[aria-label*="Plan title" i]',
      '[data-automation-id*="plan" i]'
    ];

    for (const selector of planNameSelectors) {
      const element = document.querySelector(selector);
      if (!element || !element.textContent) continue;

      const text = element.textContent.trim();
      if (this.isValidPlanName(text)) {
        this.planData.planName = text;
        return;
      }
    }

    const breadcrumbCandidates = document.querySelectorAll('[role="navigation"] a, .ms-Breadcrumb-itemLink');
    for (const candidate of breadcrumbCandidates) {
      const text = candidate.textContent?.trim();
      if (this.isValidPlanName(text)) {
        this.planData.planName = text;
        return;
      }
    }

    // Fallback: extract from URL
    const url = window.location.href;
    const urlPatterns = [
      /\/([^\/]+)(?:\?|$)/, // Last path segment
      /plan[^\/]*\/([^\/\?]+)/, // After 'plan'
      /title=([^&]+)/ // URL parameter
    ];

    for (const pattern of urlPatterns) {
      const match = url.match(pattern);
      if (!match || !match[1]) continue;

      const decoded = decodeURIComponent(match[1])
        .replace(/[-_]/g, ' ')
        .trim();
      if (this.isValidPlanName(decoded)) {
        this.planData.planName = decoded;
        return;
      }
    }

    this.planData.planName = 'Unknown Plan';
  }

  extractCurrentView() {
    // Extract current view from UI elements
    const viewSelectors = [
      '.selected [aria-label*="Grid"]',
      '.active [aria-label*="Grid"]',
      'button[aria-pressed="true"]',
      '.view-selector .selected'
    ];

    for (const selector of viewSelectors) {
      const element = document.querySelector(selector);
      if (element) {
        this.planData.currentView = element.textContent.trim() || 'Grid';
        break;
      }
    }

    // Look for Grid/Board indicators
    if (document.querySelector('[aria-label*="Grid"]')) {
      this.planData.currentView = 'Grid';
    } else if (document.querySelector('[aria-label*="Board"]')) {
      this.planData.currentView = 'Board';
    }

    // Default fallback
    if (!this.planData.currentView) {
      this.planData.currentView = 'Grid';
    }
  }

  extractAccessLevel() {
    delete this.planData.accessLevel;

    const accessSelectors = [
      '[aria-label*=" access" i]',
      '.title-side-tag',
      '[data-access-level]',
      '[class*="access"]'
    ];

    const labelMap = new Map([
      ['basic access', 'Basic access'],
      ['premium access', 'Premium access'],
      ['trial access', 'Trial access'],
      ['guest access', 'Guest access'],
      ['preview access', 'Preview access']
    ]);

    const cleanValue = (value) => {
      if (!value) return null;
      const trimmed = value.trim();
      if (!trimmed) return null;
      const simplified = trimmed
        .replace(/[^a-z0-9\s]/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      if (!simplified) return null;

      const match = simplified.match(/(basic|premium|trial|guest|preview) access/i);
      if (match) {
        const normalized = match[0].toLowerCase();
        return labelMap.get(normalized) || match[0];
      }

      return null;
    };

    for (const selector of accessSelectors) {
      const elements = document.querySelectorAll(selector);
      for (const element of elements) {
        const candidates = [
          element.getAttribute?.('aria-label'),
          element.getAttribute?.('data-access-level'),
          element.textContent
        ];

        for (const candidate of candidates) {
          const cleaned = cleanValue(candidate);
          if (cleaned) {
            this.planData.accessLevel = cleaned;
            return;
          }
        }
      }
    }
  }

  extractTaskInfo() {
    this.taskData = [];
    this.extractGridViewTasks();
    this.extractBoardViewTasks();
    this.extractTaskDetails();
  }

  extractGridViewTasks() {
    // Try multiple approaches for different Planner versions
    this.extractFromGridRows();
    this.extractFromTaskList();
    this.extractFromTableStructure();
  }

  extractFromGridRows() {
    const taskRows = document.querySelectorAll('[role="row"]');

    taskRows.forEach((row, index) => {
      if (index === 0) return; // Skip header row

      const task = {};
      const cells = row.querySelectorAll('[role="gridcell"]');

      cells.forEach(cell => {
        const ariaLabel = cell.getAttribute('aria-label');
        if (!ariaLabel) return;

        // Extract task name
        if (ariaLabel.includes('Task Name')) {
          const taskNameElement = cell.querySelector('span, div, button');
          if (taskNameElement) {
            const candidateName = taskNameElement.textContent.trim();
            if (!this.shouldSkipTaskName(candidateName)) {
              task.name = candidateName;
            }
          }
        }

        // Extract assigned user
        if (ariaLabel.includes('Assigned to')) {
          task.assignedTo = this.extractAssignedUser(cell);
        }

        // Extract progress percentage
        if (ariaLabel.includes('% complete')) {
          const percentMatch = ariaLabel.match(/(\d+)%/);
          if (percentMatch) {
            task.progress = parseInt(percentMatch[1]);
          }
        }

        // Extract priority
        if (ariaLabel.includes('Priority')) {
          const priorityMatch = ariaLabel.match(/Priority\s+(\w+)/);
          if (priorityMatch) {
            task.priority = priorityMatch[1];
          }
        }

        // Extract bucket
        if (ariaLabel.includes('Bucket')) {
          const bucketMatch = ariaLabel.match(/Bucket\s+([^.]+)/);
          if (bucketMatch) {
            task.bucket = bucketMatch[1].trim();
          }
        }

        // Extract start date
        if (ariaLabel.includes('Start')) {
          task.startDate = this.extractDate(ariaLabel);
        }

        // Extract finish date
        if (ariaLabel.includes('Finish')) {
          task.finishDate = this.extractDate(ariaLabel);
        }

        // Extract completion status
        if (ariaLabel.includes('Mark as completed')) {
          task.completed = ariaLabel.includes('checked');
        }
      });

      if (task.name) {
        if (!task.assignedTo) {
          task.assignedTo = this.findAssigneeInContainer(row);
        }
        task.id = this.generateTaskId(task.name);
        task.source = 'grid-rows';
        this.taskData.push(task);
      }
    });
  }

  extractFromTaskList() {
    // Look for tasks using aria-label patterns found in the analysis
    const taskElements = document.querySelectorAll('[aria-label*="Task Name"]');

    taskElements.forEach(element => {
      const task = {};
      const ariaLabel = element.getAttribute('aria-label');

      if (ariaLabel) {
        // Extract task name from aria-label values like "Task Name <task>. Use the space..."
        const taskNameMatch = ariaLabel.match(/Task Name\s+([^.]+)/);
        if (taskNameMatch) {
          task.name = taskNameMatch[1].trim();
          if (this.shouldSkipTaskName(task.name)) {
            task.name = null;
            return;
          }
        }

        // Extract row information
        const rowMatch = ariaLabel.match(/Row (\d+) of (\d+)/);
        if (rowMatch) {
          task.rowNumber = parseInt(rowMatch[1]);
          task.totalRows = parseInt(rowMatch[2]);
        }

        // Extract column information
        const columnMatch = ariaLabel.match(/Column (\d+) of (\d+)/);
        if (columnMatch) {
          task.columnNumber = parseInt(columnMatch[1]);
          task.totalColumns = parseInt(columnMatch[2]);
        }
      }

      // Look for assigned user in nearby elements or parent container
      const container = element.closest('[role="row"], .ms-DetailsRow, .task-container') ||
                       element.parentElement?.parentElement || element.parentElement;

      if (container) {
        const detectedAssignee = this.findAssigneeInContainer(container);
        if (detectedAssignee) {
          task.assignedTo = detectedAssignee;
        }
        // Look for completion status
        const completeButton = container.querySelector('.completeButtonIcon, [aria-label*="Mark as completed"]');
        if (completeButton) {
          // Check if task is completed (you'd need to inspect the actual state)
          const isCompleted = completeButton.closest('[aria-checked="true"]') !== null;
          task.completed = isCompleted;
          task.progress = isCompleted ? 100 : 0;
        }
      }

      if (task.name) {
        task.id = this.generateTaskId(task.name);
        task.source = 'aria-label-tasks';

        // Avoid duplicates
        const exists = this.taskData.find(t => t.name === task.name);
        if (!exists) {
          this.taskData.push(task);
        }
      }
    });

    // Also look for Quick look elements as they might contain task info
    const quickLookElements = document.querySelectorAll('[aria-label*="Quick look"]');
    quickLookElements.forEach(element => {
      const container = element.closest('[role="row"], .task-row') || element.parentElement;
      if (container) {
        const task = {};

        // Try to find task name in nearby elements
        const nameElement = container.querySelector('[aria-label*="Task Name"]');
        if (nameElement) {
          const ariaLabel = nameElement.getAttribute('aria-label');
          const taskNameMatch = ariaLabel.match(/Task Name\s+([^.]+)/);
          if (taskNameMatch) {
            task.name = taskNameMatch[1].trim();
            if (this.shouldSkipTaskName(task.name)) {
              return;
            }
            task.id = this.generateTaskId(task.name);
            task.source = 'quick-look';

            const exists = this.taskData.find(t => t.name === task.name);
            if (!exists) {
              this.taskData.push(task);
            }
          }
        }
      }
    });
  }

  extractFromTableStructure() {
    // Extract from table structure visible in screenshot
    const rows = document.querySelectorAll('tr');

    rows.forEach(row => {
      const cells = Array.from(row.querySelectorAll('td'));
      if (cells.length < 2) return;

      const task = {};
      const firstCell = cells[0];
      const taskName = firstCell.textContent.trim();

      if (!taskName || taskName === 'Task Name' || taskName.length <= 1) return;
      if (this.shouldSkipTaskName(taskName)) return;

      task.name = taskName;

      const rowAssignee = this.findAssigneeInContainer(row);
      if (rowAssignee) {
        task.assignedTo = rowAssignee;
      } else {
        for (let cellIndex = 1; cellIndex < cells.length; cellIndex += 1) {
          const cell = cells[cellIndex];
          const avatar = cell.querySelector('img[alt], img[title], .avatar, [class*="avatar"], [title]');
          if (!avatar) continue;

          const possibleValues = [
            avatar.getAttribute?.('title'),
            avatar.getAttribute?.('alt'),
            avatar.textContent
          ];

          for (const value of possibleValues) {
            const candidate = this.extractAssigneeFromText(value);
            if (candidate) {
              task.assignedTo = candidate;
              break;
            }
          }

          if (task.assignedTo) {
            break;
          }
        }
      }

      const checkbox = row.querySelector('input[type="checkbox"]');
      if (checkbox) {
        task.completed = checkbox.checked;
        task.progress = checkbox.checked ? 100 : 0;
      }

      task.id = this.generateTaskId(task.name);
      task.source = 'table-structure';

      const exists = this.taskData.find(t => t.name === task.name);
      if (!exists) {
        this.taskData.push(task);
      }
    });
  }

  extractBoardViewTasks() {
    const taskCards = document.querySelectorAll('[data-testid="task-card"], .task-card, [class*="task"], [class*="card"]');

    taskCards.forEach(card => {
      const task = {};

      const nameElement = card.querySelector('h3, h4, .task-title, [class*="title"]');
      if (nameElement) {
        task.name = nameElement.textContent.trim();
      }

      if (!task.name || this.shouldSkipTaskName(task.name)) {
        return;
      }

      const avatars = card.querySelectorAll('.avatar, [class*="avatar"], [class*="assigned"]');
      if (avatars.length > 0) {
        task.assignedTo = Array.from(avatars).map(avatar =>
          avatar.getAttribute('title') || avatar.getAttribute('alt') || 'Unknown'
        );
      }

      const labels = card.querySelectorAll('.label, [class*="label"], [class*="category"]');
      if (labels.length > 0) {
        task.labels = Array.from(labels).map(label => label.textContent.trim());
      }

      if (task.name) {
        task.id = this.generateTaskId(task.name);
        task.viewType = 'board';
        this.taskData.push(task);
      }
    });
  }

  extractTaskDetails() {
    const detailPanel = document.querySelector('[class*="detail"], [class*="panel"]');
    if (!detailPanel) return;

    const taskDetail = {};

    const titleElement = detailPanel.querySelector('h1, h2, .title, [class*="title"]');
    if (titleElement) {
      taskDetail.name = titleElement.textContent.trim();
      if (this.shouldSkipTaskName(taskDetail.name)) {
        taskDetail.name = null;
        return;
      }
    }

    const descriptionElement = detailPanel.querySelector('[class*="description"], textarea, [contenteditable]');
    if (descriptionElement) {
      taskDetail.description = descriptionElement.textContent.trim();
    }

    const checklistItems = detailPanel.querySelectorAll('[type="checkbox"]');
    if (checklistItems.length > 0) {
      taskDetail.checklist = Array.from(checklistItems).map(item => ({
        text: item.nextElementSibling ? item.nextElementSibling.textContent.trim() : '',
        completed: item.checked
      }));
    }

    const attachments = detailPanel.querySelectorAll('[class*="attachment"], .file');
    if (attachments.length > 0) {
      taskDetail.attachments = Array.from(attachments).map(attachment => ({
        name: attachment.textContent.trim(),
        url: attachment.href || ''
      }));
    }

    if (taskDetail.name) {
      taskDetail.id = this.generateTaskId(taskDetail.name);
      taskDetail.viewType = 'detail';
      this.taskData.push(taskDetail);
    }
  }

  extractBucketInfo() {
    const buckets = [];
    const bucketHeaders = document.querySelectorAll('[class*="bucket"], [class*="column"] h2, [class*="column"] h3');

    bucketHeaders.forEach(header => {
      const bucketName = header.textContent.trim();
      if (bucketName && !buckets.includes(bucketName)) {
        buckets.push(bucketName);
      }
    });

    this.planData.buckets = buckets;
  }

  extractAssignedUser(element) {
    if (!element) return null;

    const nodes = [element, ...element.querySelectorAll('[aria-label], [title], [alt], span, div')];
    for (const node of nodes) {
      const possibleValues = [
        node.getAttribute?.('aria-label'),
        node.getAttribute?.('title'),
        node.getAttribute?.('alt'),
        node.textContent
      ];

      for (const value of possibleValues) {
        const assignee = this.extractAssigneeFromText(value);
        if (assignee) {
          return assignee;
        }
      }
    }

    return null;
  }

  extractDate(text) {
    const dateMatch = text.match(/(\d{1,2}\/\d{1,2}\/\d{4}|\d{4}-\d{1,2}-\d{1,2})/);
    return dateMatch ? dateMatch[1] : null;
  }

  generateTaskId(taskName) {
    return taskName.toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50);
  }

  cleanPlanTitle(title) {
    if (!title) return null;
    const withoutProduct = title.replace(/\s+-\s+Microsoft\s+Planner.*$/i, '').trim();
    if (this.isValidPlanName(withoutProduct)) {
      return withoutProduct;
    }

    const hyphenIndex = withoutProduct.indexOf(' - ');
    if (hyphenIndex > 0) {
      const candidate = withoutProduct.slice(0, hyphenIndex).trim();
      if (this.isValidPlanName(candidate)) {
        return candidate;
      }
    }

    return null;
  }

  isValidPlanName(name) {
    if (!name) return false;
    const normalized = name.trim();
    if (normalized.length < 2) return false;

    const lowered = normalized.toLowerCase();
    const invalidNames = new Set([
      'planner',
      'microsoft planner',
      'my plans',
      'plans',
      'plan',
      'grid',
      'board',
      'charts',
      'filters'
    ]);

    if (invalidNames.has(lowered)) return false;
    if (/^view:?\s*/i.test(normalized)) return false;
    return true;
  }

  shouldSkipTaskName(name) {
    if (!name) return true;
    const normalized = name.trim().toLowerCase();
    if (!normalized) return true;

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
      return true;
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
      this.startsWithTerm(normalized, entry) ||
      this.startsWithTerm(simplified, entry)
    )) {
      return true;
    }

    if (/^\d+%/.test(normalized)) return true;

    return false;
  }

  extractAssigneeFromText(text) {
    if (!text) return null;

    const normalized = text.replace(/\s+/g, ' ').trim();
    if (!normalized) return null;

    const assignedMatch = normalized.match(/Assigned to\s+([^.,]+)/i);
    if (assignedMatch) {
      const candidate = assignedMatch[1].trim();
      if (!candidate || /unassigned/i.test(candidate) || /not assigned/i.test(candidate)) {
        return 'Unassigned';
      }
      return candidate;
    }

    if (/not assigned/i.test(normalized)) {
      return 'Unassigned';
    }

    if (/unassigned/i.test(normalized)) {
      return 'Unassigned';
    }

    if (/assign/i.test(normalized) && !/assigned/i.test(normalized)) {
      return null;
    }

    if (/^\d+%$/.test(normalized)) {
      return null;
    }

    return normalized;
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

  findAssigneeInContainer(container) {
    if (!container) return null;

    const nodes = new Set();
    const selectors = [
      '[aria-label*="Assigned" i]',
      '[title*="Assigned" i]',
      '[aria-label*="Assignee" i]',
      '[title*="Assignee" i]',
      '[data-testid*="assign" i]',
      'img[alt]',
      'img[title]',
      '.ms-Persona',
      '.persona'
    ];

    selectors.forEach(selector => {
      container.querySelectorAll(selector).forEach(node => nodes.add(node));
    });

    const row = container.closest('[role="row"]');
    if (row) {
      selectors.forEach(selector => {
        row.querySelectorAll(selector).forEach(node => nodes.add(node));
      });
    }

    for (const node of nodes) {
      const possibleValues = [
        node.getAttribute?.('aria-label'),
        node.getAttribute?.('title'),
        node.getAttribute?.('alt'),
        node.textContent
      ];

      for (const value of possibleValues) {
        const assignee = this.extractAssigneeFromText(value);
        if (assignee) {
          return assignee;
        }
      }
    }

    return null;
  }

  setupMutationObserver() {
    const observer = new MutationObserver((mutations) => {
      let shouldUpdate = false;

      mutations.forEach(mutation => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          const hasTaskElements = Array.from(mutation.addedNodes).some(node => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              return node.querySelector('[role="row"], [class*="task"], [class*="card"]');
            }
            return false;
          });

          if (hasTaskElements) {
            shouldUpdate = true;
          }
        }
      });

      if (shouldUpdate) {
        clearTimeout(this.updateTimeout);
        this.updateTimeout = setTimeout(() => {
          this.extractData();
        }, 1000);
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true
    });

    this.observers.push(observer);
  }

  async sendDataToBackground() {
    const extractedData = {
      url: window.location.href,
      timestamp: new Date().toISOString(),
      planData: this.planData,
      taskData: this.taskData,
      extractionSuccess: this.taskData.length > 0 || Object.keys(this.planData).length > 0
    };

    try {
      await chrome.runtime.sendMessage({
        action: 'plannerDataExtracted',
        data: extractedData
      });
    } catch (error) {
      console.log('Extension context invalid, storing locally:', error);
      localStorage.setItem('plannerData', JSON.stringify(extractedData));
    }

    sessionStorage.setItem('currentPlannerData', JSON.stringify(extractedData));
  }

  getCurrentData() {
    return {
      planData: this.planData,
      taskData: this.taskData
    };
  }

  manualExtract() {
    this.extractData();
    return this.getCurrentData();
  }

  destroy() {
    this.observers.forEach(observer => observer.disconnect());
    this.observers = [];
    clearTimeout(this.updateTimeout);
  }
}

// Initialize the extractor
let plannerExtractor;

// Make sure we only initialize once
if (!window.plannerExtractorInitialized) {
  plannerExtractor = new PlannerDataExtractor();
  plannerExtractor.init();
  window.plannerExtractorInitialized = true;

  // Make extractor available globally for debugging
  window.plannerExtractor = plannerExtractor;
}

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'extractPlannerData') {
    const data = plannerExtractor.manualExtract();
    sendResponse({ success: true, data });
    return true;
  }

  if (request.action === 'getCurrentData') {
    const data = plannerExtractor.getCurrentData();
    sendResponse({ success: true, data });
    return true;
  }
});

console.log('Microsoft Planner Content Script loaded');
