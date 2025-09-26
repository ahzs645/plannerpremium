/**
 * Microsoft Planner Data Extractor Module
 * ES6 Module for extracting data from Planner interface
 */

export class PlannerDataExtractor {
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
    // Extract plan name from breadcrumb
    const planNameElement = document.querySelector('.ms-Breadcrumb-itemLink');
    if (planNameElement) {
      this.planData.planName = planNameElement.textContent.trim();
    }

    // Extract project title from page
    const projectTitleElements = document.querySelectorAll('[aria-label*="Project"]');
    projectTitleElements.forEach(el => {
      if (el.textContent && el.textContent.trim()) {
        this.planData.projectTitle = el.textContent.trim();
      }
    });

    // Extract current view
    const viewElements = document.querySelectorAll('[aria-label*="View"]');
    viewElements.forEach(el => {
      if (el.classList.contains('selected') || el.getAttribute('aria-selected') === 'true') {
        this.planData.currentView = el.textContent.trim();
      }
    });
  }

  extractTaskInfo() {
    this.taskData = [];
    this.extractGridViewTasks();
    this.extractBoardViewTasks();
    this.extractTaskDetails();
  }

  extractGridViewTasks() {
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
            task.name = taskNameElement.textContent.trim();
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
        task.id = this.generateTaskId(task.name);
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
    const userElement = element.querySelector('[title], [alt], span, div');
    if (userElement) {
      return userElement.getAttribute('title') ||
             userElement.getAttribute('alt') ||
             userElement.textContent.trim();
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