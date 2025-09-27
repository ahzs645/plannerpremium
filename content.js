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
    delete this.planData.currentView;

    const activeToggle = document.querySelector(
      '[aria-pressed="true"][aria-label], [role="tab"][aria-selected="true"]'
    );

    if (activeToggle) {
      const label = activeToggle.getAttribute('aria-label') || activeToggle.textContent;
      if (label) {
        this.planData.currentView = label.trim();
      }
    }

    if (!this.planData.currentView) {
      const url = window.location.href.toLowerCase();
      if (url.includes('/view/board')) {
        this.planData.currentView = 'Board';
      } else if (url.includes('/view/grid')) {
        this.planData.currentView = 'Grid';
      }
    }

    if (!this.planData.currentView) {
      const boardIndicators = [
        document.querySelector('li.board-column, [data-is-focusable].board-column'),
        document.querySelector('[data-dnd-role="card"], .task-board-card')
      ];

      if (boardIndicators.some(Boolean)) {
        this.planData.currentView = 'Board';
      }
    }

    if (!this.planData.currentView) {
      const gridIndicators = [
        document.querySelector('[role="grid"], .task-grid-view'),
        document.querySelector('[data-automation-id*="grid"]'),
        document.querySelector('[role="rowgroup"]')
      ];

      if (gridIndicators.some(Boolean)) {
        this.planData.currentView = 'Grid';
      }
    }

    if (this.planData.currentView) {
      const normalized = this.planData.currentView.toLowerCase();
      if (normalized.includes('board')) {
        this.planData.currentView = 'Board';
      } else if (normalized.includes('grid')) {
        this.planData.currentView = 'Grid';
      }
    }

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
    if (this.isBoardView()) {
      return;
    }

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
            const candidateName = this.cleanTaskTitle(taskNameElement.textContent);
            if (candidateName && !this.shouldSkipTaskName(candidateName)) {
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
          task.name = this.cleanTaskTitle(taskNameMatch[1]);
          if (!task.name || this.shouldSkipTaskName(task.name)) {
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
            task.name = this.cleanTaskTitle(taskNameMatch[1]);
            if (!task.name || this.shouldSkipTaskName(task.name)) {
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
      const taskName = this.cleanTaskTitle(firstCell.textContent);

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
    const columns = document.querySelectorAll('li.board-column, [data-is-focusable].board-column');

    if (!columns.length) {
      const fallbackCards = document.querySelectorAll('[data-dnd-role="card"], .task-board-card');
      fallbackCards.forEach(card => {
        const task = this.buildBoardTaskFromCard(card);
        if (task) {
          this.taskData.push(task);
        }
      });
      return;
    }

    columns.forEach(column => {
      const bucketName = this.getBucketNameFromColumn(column);
      const cards = column.querySelectorAll('[data-dnd-role="card"], .task-board-card');

      cards.forEach(card => {
        const task = this.buildBoardTaskFromCard(card, bucketName);
        if (task) {
          this.taskData.push(task);
        }
      });
    });
  }

  extractTaskDetails() {
    const detailPanel = document.querySelector('[class*="detail"], [class*="panel"]');
    if (!detailPanel) return;

    const taskDetail = {};

    const titleElement = detailPanel.querySelector('h1, h2, .title, [class*="title"]');
    if (titleElement) {
      taskDetail.name = this.cleanTaskTitle(titleElement.textContent);
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
    const boardColumns = document.querySelectorAll('li.board-column, [data-is-focusable].board-column');
    boardColumns.forEach(column => {
      const bucketName = this.getBucketNameFromColumn(column);
      const cleanedName = this.cleanTaskTitle(bucketName);
      const lowered = cleanedName?.toLowerCase();
      if (cleanedName && lowered !== 'add bucket' && lowered !== 'add new bucket' && !buckets.includes(cleanedName)) {
        buckets.push(cleanedName);
      }
    });

    if (buckets.length === 0) {
      const bucketHeaders = document.querySelectorAll('[class*="bucket"], [class*="column"] h2, [class*="column"] h3');

      bucketHeaders.forEach(header => {
        const bucketName = this.cleanTaskTitle(header.textContent);
        const lowered = bucketName?.toLowerCase();
        if (bucketName && lowered !== 'add bucket' && lowered !== 'add new bucket' && !buckets.includes(bucketName)) {
          buckets.push(bucketName);
        }
      });
    }

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
      'add bucket',
      'add new bucket',
      'add task',
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
      'bucket',
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

    if (/assign this task/i.test(normalized)) {
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

  normalizeTaskName(name) {
    return name ? name.trim().toLowerCase() : '';
  }

  async createTaskAndOpenDetails(taskName, options = {}) {
    const { openDetails = true } = options;
    const createdTask = await this.addTask(taskName, options);
    if (!openDetails) {
      return createdTask;
    }
    await this.openTaskDetails(taskName);
    return this.findTaskDataByName(this.normalizeTaskName(taskName)) || createdTask;
  }

  async addTask(taskName, options = {}) {
    const normalized = this.normalizeTaskName(taskName);
    if (!normalized) {
      throw new Error('Task name is required');
    }

    const existingRow = this.findTaskRowElementByName(normalized);
    if (existingRow) {
      return this.findTaskDataByName(normalized);
    }

    if (this.isBoardView()) {
      return this.addTaskInBoard(taskName, normalized, options);
    }

    return this.addTaskInGrid(taskName, normalized);
  }

  async addTaskInGrid(taskName, normalizedName) {
    const addControl = this.findAddTaskControl();
    if (!addControl) {
      const inlineInput = document.querySelector('input[aria-label="Add New Row"], textarea[aria-label="Add new task" i]');
      if (!inlineInput) {
        throw new Error('Add new task control not found');
      }
      inlineInput.focus({ preventScroll: true });
      inlineInput.value = '';
      inlineInput.dispatchEvent(new Event('input', { bubbles: true }));
      inlineInput.value = taskName;
      inlineInput.dispatchEvent(new Event('input', { bubbles: true }));
      inlineInput.dispatchEvent(new Event('change', { bubbles: true }));
      this.dispatchKeyboardSequence(inlineInput, ['Enter']);
    } else {
      this.simulateClick(addControl);

      const input = await this.waitForElement([
        'input[aria-label="Add New Row"]',
        'input[aria-label*="Add new task" i]',
        'textarea[aria-label*="Add new task" i]'
      ], { timeout: 6000 }).catch(() => null);

      if (!input) {
        throw new Error('Add task input did not appear');
      }

      input.focus({ preventScroll: true });
      input.value = '';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.value = taskName;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      if (typeof input.blur === 'function') {
        input.blur();
      }

      this.dispatchKeyboardSequence(input, ['Enter']);
    }

    await this.waitForCondition(() => this.findTaskRowElementByName(normalizedName), { timeout: 10000 }).catch(() => {
      throw new Error('Task row was not created');
    });

    await this.waitForCondition(() => this.findTaskDataByName(normalizedName), { timeout: 8000 }).catch(() => null);

    this.extractData();

    return this.findTaskDataByName(normalizedName);
  }

  async addTaskInBoard(taskName, normalizedName, options = {}) {
    const targetBucket = options.bucketName ? this.normalizeTaskName(options.bucketName) : null;

    let columnForTask = null;
    if (targetBucket) {
      const columns = document.querySelectorAll('li.board-column, [data-is-focusable].board-column');
      columnForTask = Array.from(columns).find(column => {
        const name = this.normalizeTaskName(this.getBucketNameFromColumn(column));
        return name && name === targetBucket;
      }) || null;
    }

    let addControl = columnForTask
      ? this.findBoardAddTaskButton(columnForTask)
      : this.findBoardAddTaskButton();

    if (!addControl && !document.querySelector('.add-task-card')) {
      throw new Error('Add task button not found in board view');
    }

    if (!document.querySelector('.add-task-card') && addControl) {
      this.simulateClick(addControl);
    }

    const addCard = await this.waitForElement('.add-task-card', { timeout: 5000 });
    if (!addCard) {
      throw new Error('Add task card did not appear');
    }

    const nameInput = addCard.querySelector('input[data-cy-new-task-name], input[placeholder*="task" i]');
    if (!nameInput) {
      throw new Error('Add task input not found in board view');
    }

    nameInput.focus({ preventScroll: true });
    nameInput.value = '';
    nameInput.dispatchEvent(new Event('input', { bubbles: true }));
    nameInput.value = taskName;
    nameInput.dispatchEvent(new Event('input', { bubbles: true }));
    nameInput.dispatchEvent(new Event('change', { bubbles: true }));

    const confirmButton = addCard.querySelector('.add-task-card button.bottom-add-task-button, .add-task-card button[aria-label="Add task"], .add-task-card button.ms-Button--primary');
    if (confirmButton) {
      this.simulateClick(confirmButton);
    } else {
      this.dispatchKeyboardSequence(nameInput, ['Enter']);
    }

    await this.waitForCondition(() => !document.querySelector('.add-task-card'), { timeout: 6000 }).catch(() => null);

    const newCard = await this.waitForCondition(() => this.findTaskRowElementByName(normalizedName), { timeout: 6000 }).catch(() => null);
    if (!newCard) {
      throw new Error('Task card was not created in board view');
    }

    this.extractData();

    return this.findTaskDataByName(normalizedName);
  }

  async openTaskDetails(taskName) {
    const normalized = this.normalizeTaskName(taskName);
    const row = await this.waitForCondition(() => this.findTaskRowElementByName(normalized), { timeout: 5000 });

    if (!row) {
      throw new Error('Task row not found');
    }

    if (this.isBoardView()) {
      await this.activateBoardCard(row);
    } else {
      this.simulateHover(row);

      const detailsButton = await this.waitForCondition(() =>
        row.querySelector('[aria-label*="Open details" i], [title*="Open details" i]')
      , { timeout: 2000 }).catch(() => null);

      if (detailsButton) {
        this.simulateClick(detailsButton);
      } else {
        this.simulateClick(row);
        this.simulateDoubleClick(row);
        this.dispatchKeyboardSequence(row, ['Enter']);
      }
    }

    const detailPanel = await this.waitForCondition(() => this.getTaskDetailPanelElement(), { timeout: 12000 }).catch(() => null);

    if (!detailPanel) {
      throw new Error('Task detail panel did not open');
    }

    if (this.isBoardView()) {
      await this.populateBoardTaskDetails(detailPanel);
    }

    this.extractTaskDetails();
    this.mergeTaskDetail(normalized);
    this.sendDataToBackground();

    return this.findTaskDataByName(normalized);
  }

  async activateBoardCard(card) {
    if (!card) return;

    this.simulateHover(card);
    this.simulateClick(card);

    const immediatePanel = await this.waitForCondition(() => this.getTaskDetailPanelElement(), { timeout: 1200 }).catch(() => null);
    if (immediatePanel) {
      return;
    }

    const quickAction = await this.waitForCondition(() =>
      card.querySelector('[aria-label*="Open details" i], [title*="Open details" i], button[data-automation-id*="details" i]')
    , { timeout: 1500 }).catch(() => null);

    if (quickAction) {
      this.simulateClick(quickAction);
      return;
    }

    if (typeof card.focus === 'function') {
      card.focus({ preventScroll: true });
    }

    this.dispatchKeyboardSequence(card, ['Enter']);
    this.simulateDoubleClick(card);
  }

  getTaskDetailPanelElement() {
    const selectors = [
      '.taskDetailsPane',
      '.taskDetailsContainer',
      '.task-details-pane',
      '.ms-Panel.taskDetailsPanel',
      '.ms-Modal.ms-Dialog.taskDetailsContainer',
      '[data-automation-id*="taskdetails" i]',
      '[class*="taskDetails" i]',
      '[role="dialog"][aria-label*="task" i]',
      '[role="dialog"][aria-labelledby*="task" i]'
    ];

    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (element) {
        return element;
      }
    }

    const dialogs = document.querySelectorAll('div[role="dialog"], section[role="dialog"], section[role="region"]');
    for (const dialog of dialogs) {
      const label = `${dialog.getAttribute('aria-label') || ''} ${dialog.textContent || ''}`.toLowerCase();
      if (label.includes('task') && label.includes('detail')) {
        return dialog;
      }
    }

    return null;
  }

  findAddTaskControl() {
    const selectors = [
      '#data-cy-new-row',
      '[data-cy="new-row"]',
      '[id^="data-cy-new-row"]',
      '.new-row-placeholder',
      '[aria-label="Add new task"]',
      '[aria-label*="Add new task" i]',
      'button[title="Add task"]',
      'button[data-automation-id*="addtask" i]',
      'button[aria-label*="Add task" i]'
    ];

    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (element) {
        return element;
      }
    }

    return null;
  }

  findBoardAddTaskButton(context) {
    const searchRoots = [];
    if (context) {
      searchRoots.push(context);
    } else {
      const columns = document.querySelectorAll('li.board-column, [data-is-focusable].board-column');
      if (columns.length) {
        searchRoots.push(...columns);
      }
      searchRoots.push(document);
    }

    const isAddTaskButton = (button) => {
      if (!button || button.closest('.add-task-card')) return false;
      const aria = (button.getAttribute('aria-label') || '').toLowerCase();
      const title = (button.getAttribute('title') || '').toLowerCase();
      const text = (button.textContent || '').toLowerCase();
      return aria.includes('add task') || title.includes('add task') || text.includes('add task');
    };

    for (const root of searchRoots) {
      const buttons = root.querySelectorAll('button, div[role="button"], span[role="button"]');
      for (const button of buttons) {
        if (isAddTaskButton(button)) {
          return button;
        }
      }
    }

    return null;
  }

  isBoardView() {
    if (this.planData?.currentView && this.planData.currentView.toLowerCase().includes('board')) {
      return true;
    }
    return Boolean(document.querySelector('[data-dnd-role="card"], .task-board-card'));
  }

  getBucketNameFromColumn(column) {
    if (!column) return null;

    const ariaLabel = column.getAttribute('aria-label');
    if (ariaLabel) {
      const match = ariaLabel.match(/Bucket:\s*([^,.]+)/i);
      if (match && match[1]) {
        const label = this.cleanTaskTitle(match[1]);
        if (label) {
          return label;
        }
      }
    }

    const header = column.querySelector('[role="heading"], h2, h3, .column-header, .bucket-header, .bucketHeader');
    if (header) {
      const text = this.cleanTaskTitle(header.textContent);
      if (text) {
        return text;
      }
    }

    return null;
  }

  getTaskNameFromAriaLabel(label) {
    if (!label) return null;
    const match = label.match(/Task(?: Name)?\s*([^.,]+)/i);
    if (match && match[1]) {
      return match[1].trim();
    }
    return null;
  }

  cleanTaskTitle(name) {
    if (!name) return '';

    return String(name)
      .replace(/^Task Name\s+/i, '')
      .replace(/Use the space or enter key[\s\S]*$/i, '')
      .replace(/Use arrow keys[\s\S]*$/i, '')
      .replace(/Finish[\s\S]*$/i, '')
      .replace(/Assign this task/gi, '')
      .replace(/(?:day|due)\s*\d{1,2}\/\d{1,2}/gi, '')
      .replace(/\b[0-9]+%\s+complete\b/gi, '')
      .replace(/[\uE000-\uF8FF]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  buildBoardTaskFromCard(card, bucketName) {
    if (!card || card.closest('.add-task-card')) {
      return null;
    }

    const titleElement = card.querySelector('[aria-label*="Task Name"], [aria-label*="Task" i], [data-task-title], .title, .task-title, .task-name, h3, h4, [role="textbox"]');
    let name = this.cleanTaskTitle(titleElement?.textContent?.trim());

    if (!name) {
      const ariaLabel = card.getAttribute('aria-label') || titleElement?.getAttribute?.('aria-label');
      name = this.cleanTaskTitle(this.getTaskNameFromAriaLabel(ariaLabel));
    }

    if (!name || this.shouldSkipTaskName(name)) {
      return null;
    }

    const task = {
      name,
      id: this.generateTaskId(name),
      viewType: 'board'
    };

    if (bucketName) {
      task.bucket = bucketName;
    }

    const avatars = card.querySelectorAll('.avatar, [class*="avatar"], [class*="assigned"], .ms-Facepile-itemButton');
    if (avatars.length > 0) {
      const assignees = Array.from(avatars)
        .map(avatar =>
          this.extractAssigneeFromText(
            avatar.getAttribute('title') ||
            avatar.getAttribute('alt') ||
            avatar.textContent?.trim()
          )
        )
        .filter(Boolean);

      const uniqueAssignees = Array.from(new Set(assignees));
      const hasNamedAssignee = uniqueAssignees.some(name => name && name !== 'Unassigned');
      const filteredAssignees = hasNamedAssignee
        ? uniqueAssignees.filter(name => name !== 'Unassigned')
        : uniqueAssignees;

      if (filteredAssignees.length === 1) {
        task.assignedTo = filteredAssignees[0];
      } else if (filteredAssignees.length > 1) {
        task.assignedTo = filteredAssignees;
      } else {
        task.assignedTo = 'Unassigned';
      }
    }

    const labels = card.querySelectorAll('.label, [class*="label"], [class*="category"]');
    if (labels.length > 0) {
      const labelValues = Array.from(labels).map(label => label.textContent.trim()).filter(Boolean);
      if (labelValues.length > 0) {
        task.labels = labelValues;
      }
    }

    if (!task.assignedTo) {
      task.assignedTo = 'Unassigned';
    }

    return task;
  }

  findTaskRowElementByName(normalizedName) {
    if (!normalizedName) return null;

    const gridCandidates = document.querySelectorAll('[aria-label*="Task Name"], [role="gridcell"]');
    for (const candidate of gridCandidates) {
      const label = candidate.getAttribute('aria-label');
      const text = this.cleanTaskTitle(candidate.textContent);
      if (label && this.normalizeTaskName(label).includes(`task name ${normalizedName}`)) {
        return candidate.closest('[role="row"]') || candidate;
      }

      if (text && this.normalizeTaskName(text) === normalizedName) {
        return candidate.closest('[role="row"]') || candidate;
      }
    }

    const boardCards = document.querySelectorAll('[data-dnd-role="card"], .task-board-card');
    for (const card of boardCards) {
      if (card.closest('.add-task-card')) {
        continue;
      }
      const titleElement = card.querySelector('[data-task-title], .title, .task-name, [class*="title"], [role="textbox"]');
      const text = this.cleanTaskTitle(titleElement?.textContent);
      if (text && this.normalizeTaskName(text) === normalizedName) {
        return card;
      }

      const ariaLabel = card.getAttribute('aria-label') || titleElement?.getAttribute?.('aria-label');
      const labelName = this.cleanTaskTitle(this.getTaskNameFromAriaLabel(ariaLabel));
      if (labelName && this.normalizeTaskName(labelName) === normalizedName) {
        return card;
      }
    }

    return null;
  }

  findTaskDataByName(normalizedName) {
    if (!normalizedName) return null;
    return this.taskData.find(task => this.normalizeTaskName(task.name) === normalizedName) || null;
  }

  mergeTaskDetail(normalizedName) {
    if (!normalizedName) return;
    const detailIndex = this.taskData.findIndex(task =>
      task?.viewType === 'detail' && this.normalizeTaskName(task.name) === normalizedName
    );

    if (detailIndex === -1) {
      return;
    }

    const detail = this.taskData[detailIndex];
    const base = this.taskData.find((task, index) =>
      index !== detailIndex && this.normalizeTaskName(task.name) === normalizedName
    );

    if (base) {
      Object.assign(base, detail);
      this.taskData.splice(detailIndex, 1);
    }
  }

  simulateClick(element) {
    if (!element) return;

    const fire = (Ctor, type, init = {}) => {
      try {
        const event = new Ctor(type, { bubbles: true, cancelable: true, ...init });
        element.dispatchEvent(event);
      } catch (error) {
        const fallback = new Event(type, { bubbles: true, cancelable: true });
        element.dispatchEvent(fallback);
      }
    };

    if (typeof window.PointerEvent === 'function') {
      fire(window.PointerEvent, 'pointerdown');
    }
    fire(MouseEvent, 'mousedown');

    if (typeof element.focus === 'function') {
      element.focus({ preventScroll: true });
    }

    if (typeof window.PointerEvent === 'function') {
      fire(window.PointerEvent, 'pointerup');
    }
    fire(MouseEvent, 'mouseup');
    fire(MouseEvent, 'click');
  }

  simulateDoubleClick(element) {
    if (!element) return;
    const dblClick = new MouseEvent('dblclick', { bubbles: true, cancelable: true });
    element.dispatchEvent(dblClick);
  }

  async populateBoardTaskDetails(detailPanel) {
    if (!detailPanel) return;

    const formatDate = (date) => {
      const pad = (value) => String(value).padStart(2, '0');
      const month = pad(date.getMonth() + 1);
      const day = pad(date.getDate());
      const year = date.getFullYear();
      return `${month}/${day}/${year}`;
    };

    const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    const fireInputEvent = (element, type = 'input') => {
      if (typeof window !== 'undefined' && typeof window.InputEvent === 'function') {
        try {
          element.dispatchEvent(new window.InputEvent(type, { bubbles: true }));
          return;
        } catch (error) {
          // Fall back to generic event
        }
      }
      element.dispatchEvent(new Event(type, { bubbles: true }));
    };

    const dispatchInputEvents = (element) => {
      fireInputEvent(element, 'input');
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('blur', { bubbles: true }));
    };

    const typeText = async (element, text) => {
      if (!element) return;
      element.focus?.({ preventScroll: true });
      element.value = '';
      fireInputEvent(element, 'input');
      for (const char of text) {
        element.value += char;
        fireInputEvent(element, 'input');
        await wait(40);
      }
      element.dispatchEvent(new Event('change', { bubbles: true }));
    };

    const setInputValue = (selector, value) => {
      const input = detailPanel.querySelector(selector);
      if (!input) return;
      input.focus?.({ preventScroll: true });
      input.value = value;
      dispatchInputEvents(input);
    };

    const notesEditor = detailPanel.querySelector('.notes-editor[contenteditable="true"]');
    const notesText = notesEditor?.textContent || '';
    if (notesEditor && (notesText.trim() === '' || /add a note/i.test(notesText))) {
      const noteText = 'Added automatically by Planner Interface.';
      notesEditor.focus?.({ preventScroll: true });
      notesEditor.textContent = noteText;
      fireInputEvent(notesEditor, 'input');
      notesEditor.dispatchEvent(new Event('blur', { bubbles: true }));
    }

    const today = new Date();
    const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
    setInputValue('input[aria-label^="Start" i]', formatDate(today));
    setInputValue('input[aria-label*="Finish" i]', formatDate(tomorrow));

    const percentInput = detailPanel.querySelector('.ms-spinButton-input');
    if (percentInput) {
      percentInput.focus?.({ preventScroll: true });
      percentInput.value = '10';
      dispatchInputEvents(percentInput);
    }

    const checklistInput = detailPanel.querySelector('[placeholder="Add an item"], [aria-label="Add a checklist item"]');
    if (checklistInput && !checklistInput.value) {
      checklistInput.focus?.({ preventScroll: true });
      checklistInput.value = 'First checklist item';
      fireInputEvent(checklistInput, 'input');
      this.dispatchKeyboardSequence(checklistInput, ['Enter']);
      checklistInput.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const trySelectDropdown = (labelText) => {
      const label = Array.from(detailPanel.querySelectorAll('label')).find(node =>
        node.textContent?.trim().toLowerCase() === labelText
      );
      if (!label) return;
      const comboboxId = label.getAttribute('for');
      const combobox = comboboxId
        ? document.getElementById(comboboxId)
        : label.parentElement?.querySelector('[role="combobox"]');
      if (!combobox) return;

      this.simulateClick(combobox);

      const optionContainer = document.querySelector('.ms-Callout, .ms-ContextualMenu, [role="listbox"]');
      if (!optionContainer) {
        this.simulateClick(combobox);
        return;
      }

      const option = optionContainer.querySelector('[role="option"]');
      if (option) {
        this.simulateClick(option);
      }

      this.simulateClick(combobox);
    };

    trySelectDropdown('bucket');
    trySelectDropdown('priority');

    const assignButton = detailPanel.querySelector('[aria-label*="Assign this task" i], .ms-Facepile-addButton');
    if (assignButton) {
      this.simulateClick(assignButton);

      const assignCallout = await this.waitForCondition(() => {
        const callouts = Array.from(document.querySelectorAll('.ms-Callout, .ms-ContextualMenu'));
        return callouts.find(callout => callout.querySelector('input[aria-label*="enter text to search" i], input[data-automation-id*="people" i], input[placeholder*="Search" i]')) || null;
      }, { timeout: 3000 }).catch(() => null);

      const searchInput = assignCallout?.querySelector('input[aria-label*="enter text to search" i], input[data-automation-id*="people" i], input[placeholder*="Search" i]');

      if (searchInput) {
        await typeText(searchInput, 'ahmad.jalil@northernhealth.ca');
        await wait(600);

        const suggestion = await this.waitForCondition(() => {
          const candidates = Array.from(assignCallout?.querySelectorAll('[role="option"], .ms-Suggestions-item, button, .resourcePickerPersona') || []);
          return candidates.find(node => {
            const label = (node.getAttribute?.('aria-label') || node.textContent || '').toLowerCase();
            return label.includes('ahmad') || label.includes('northernhealth') || label.includes('jalil');
          }) || candidates[0] || null;
        }, { timeout: 4000 }).catch(() => null);

        if (suggestion) {
          this.simulateClick(suggestion);
          await wait(400);
        } else {
          this.dispatchKeyboardSequence(searchInput, ['Enter']);
          await wait(300);
        }
        await wait(1500);
      }
    }

    const labelWrapper = detailPanel.querySelector('.labelPickerWrapper');
    const labelTrigger = labelWrapper?.querySelector('.addLabel, button, [role="button"]');
    if (labelWrapper && labelTrigger) {
      await wait(750);
      this.simulateClick(labelTrigger);

      const labelInput = await this.waitForCondition(() =>
        labelWrapper.querySelector('input[aria-label*="label" i]')
      , { timeout: 2000 }).catch(() => null);

      if (labelInput) {
        labelInput.focus?.({ preventScroll: true });
        fireInputEvent(labelInput, 'focus');
        await wait(250);
        labelInput.value = 'red';
        fireInputEvent(labelInput, 'input');
        labelInput.dispatchEvent(new Event('change', { bubbles: true }));
        await wait(400);
        this.dispatchKeyboardSequence(labelInput, ['Enter']);
        await wait(400);
      }

      const labelMenu = await this.waitForCondition(() =>
        document.querySelector('.labelPickerSuggestionsDropdown, .ms-Callout .ms-Suggestions')
      , { timeout: 2000 }).catch(() => null);

      const gatherCandidates = (root) => root
        ? Array.from(root.querySelectorAll('.ms-Suggestions-itemButton, [role="menuitem"], [role="option"], [role="button"], button, .editableLabelWrapper'))
        : [];

      let labelCandidates = gatherCandidates(labelMenu);
      if (!labelCandidates.length) {
        labelCandidates = gatherCandidates(labelWrapper);
      }

      if (labelCandidates.length) {
        const findRedLabel = () => labelCandidates.find(node => {
          const aria = (node.getAttribute('aria-label') || '').toLowerCase();
          const text = (node.textContent || '').toLowerCase();
          const dataColor = (node.getAttribute('data-color') || node.getAttribute('data-colorid') || node.dataset?.color || '').toLowerCase();
          return aria.includes('red') || text.includes('red') || ['0', 'red', 'label1'].includes(dataColor);
        });

        const normalizeTarget = (element) => {
          if (!element) return null;
          if (element.matches('.ms-Suggestions-itemButton, [role="button"], button')) {
            return element;
          }
          return element.querySelector('.ms-Suggestions-itemButton, [role="button"], button') || element;
        };

        const targetLabel = normalizeTarget(findRedLabel()) || normalizeTarget(labelCandidates[0]) || null;
        if (targetLabel) {
          this.simulateClick(targetLabel);
          await wait(200);
        }
      }
    }
  }

  simulateHover(element) {
    if (!element) return;
    const mouseEnter = new MouseEvent('mouseenter', { bubbles: false, cancelable: true });
    element.dispatchEvent(mouseEnter);
    const mouseOver = new MouseEvent('mouseover', { bubbles: true, cancelable: true });
    element.dispatchEvent(mouseOver);
  }

  dispatchKeyboardSequence(target, keys) {
    if (!target || !keys) return;

    keys.forEach(key => {
      const options = {
        key,
        code: key,
        keyCode: key === 'Enter' ? 13 : undefined,
        which: key === 'Enter' ? 13 : undefined,
        bubbles: true,
        cancelable: true
      };

      target.dispatchEvent(new KeyboardEvent('keydown', options));
      target.dispatchEvent(new KeyboardEvent('keypress', options));
      target.dispatchEvent(new KeyboardEvent('keyup', options));
    });
  }

  waitForElement(selectors, options = {}) {
    const selectorList = Array.isArray(selectors) ? selectors.join(',') : selectors;
    return this.waitForCondition(() => document.querySelector(selectorList), options);
  }

  waitForCondition(predicate, { timeout = 5000, interval = 100 } = {}) {
    return new Promise((resolve, reject) => {
      const start = Date.now();
      let hasLoggedError = false;

      const check = () => {
        try {
          const result = predicate();
          if (result) {
            clearInterval(timer);
            resolve(result);
            return;
          }
        } catch (error) {
          if (!hasLoggedError) {
            console.warn('Condition check failed:', error);
            hasLoggedError = true;
          }
        }

        if (Date.now() - start >= timeout) {
          clearInterval(timer);
          reject(new Error('Timed out waiting for condition'));
        }
      };

      const timer = setInterval(check, interval);
      check();
    });
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
  window.plannerAutomation = {
    addTask: (name, options) => plannerExtractor.addTask(name, options),
    openTaskDetails: (name) => plannerExtractor.openTaskDetails(name),
    createTaskAndOpenDetails: (name, options) => plannerExtractor.createTaskAndOpenDetails(name, options)
  };
}

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'createPlannerTask') {
    const { taskName } = request;
    if (!taskName || !taskName.trim()) {
      sendResponse({ success: false, error: 'Task name is required' });
      return true;
    }

    plannerExtractor.addTask(taskName.trim())
      .then(task => {
        sendResponse({ success: true, task, data: plannerExtractor.getCurrentData() });
      })
      .catch(error => {
        console.error('Failed to add task:', error);
        sendResponse({ success: false, error: error.message });
      });

    return true;
  }

  if (request.action === 'openPlannerTaskDetails') {
    const { taskName } = request;
    if (!taskName || !taskName.trim()) {
      sendResponse({ success: false, error: 'Task name is required' });
      return true;
    }

    plannerExtractor.openTaskDetails(taskName.trim())
      .then(task => {
        sendResponse({ success: true, task, data: plannerExtractor.getCurrentData() });
      })
      .catch(error => {
        console.error('Failed to open task details:', error);
        sendResponse({ success: false, error: error.message });
      });

    return true;
  }

  if (request.action === 'createTaskAndOpenDetails') {
    const { taskName, options } = request;
    if (!taskName || !taskName.trim()) {
      sendResponse({ success: false, error: 'Task name is required' });
      return true;
    }

    plannerExtractor.createTaskAndOpenDetails(taskName.trim(), options)
      .then(task => {
        sendResponse({ success: true, task, data: plannerExtractor.getCurrentData() });
      })
      .catch(error => {
        console.error('Failed to create task and open details:', error);
        sendResponse({ success: false, error: error.message });
      });

    return true;
  }

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
