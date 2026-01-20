/**
 * Detailed Extractor for Planner Exporter
 * Opens each task's detail panel to extract full information
 * Best for: Complete exports including checklists, notes, descriptions
 */

import { DomUtils } from './domUtils.js';
import { BasicExtractor } from './basicExtractor.js';

export const DetailedExtractor = {
  // Panel selectors
  SELECTORS: {
    detailsButton: 'i[aria-label*="Open details"], button[aria-label*="Open details"]',
    closeButton: 'button[aria-label*="Close"], button i[data-icon-name="ChromeClose"], button[aria-label="Dismiss"]',
    detailsPanel: '[class*="TaskDetailsPanel"], [class*="taskDetailsPanel"], [role="dialog"], [class*="Panel"][class*="isOpen"]',

    // Panel content selectors
    taskName: '[class*="taskName"] input, [class*="TaskName"] input, [aria-label="Task name"]',
    startDate: '[aria-label*="Start"], [class*="startDate"] input',
    dueDate: '[aria-label*="Due"], [aria-label*="Finish"], [class*="dueDate"] input',
    percentComplete: '[aria-label*="Complete"], [aria-label*="Progress"]',
    priority: '[aria-label*="Priority"]',
    bucket: '[aria-label*="Bucket"], [class*="bucket"]',
    description: '[aria-label*="Notes"], [aria-label*="Description"], [class*="description"] textarea',
    checklist: '[class*="checklist"], [class*="Checklist"], [role="list"]',
    checklistItem: '[class*="checklistItem"], [role="listitem"]',
    assignedTo: '[class*="persona"], [class*="Persona"], [class*="assignee"]'
  },

  /**
   * Open the details panel for a task row
   */
  async openTaskDetails(row) {
    // Find the details button in the row
    const nameCell = row.querySelector('[role="treeitem"]');
    if (!nameCell) return false;

    // Try different selectors for the info/details button
    const buttonSelectors = [
      'i[aria-label*="Open details"]',
      'button[aria-label*="Open details"]',
      'i[data-icon-name="Info"]',
      'button[title*="details"]'
    ];

    let button = null;
    for (const selector of buttonSelectors) {
      const icon = nameCell.querySelector(selector);
      if (icon) {
        button = icon.closest('button') || icon;
        break;
      }
    }

    if (!button) {
      // Try hovering over the row to reveal the button
      row.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
      await DomUtils.sleep(200);

      for (const selector of buttonSelectors) {
        const icon = nameCell.querySelector(selector);
        if (icon) {
          button = icon.closest('button') || icon;
          break;
        }
      }
    }

    if (!button) return false;

    // Click the button
    button.click();

    // Wait for panel to open
    try {
      await DomUtils.waitForElement(this.SELECTORS.detailsPanel, 3000);
      await DomUtils.sleep(300); // Extra wait for content to load
      return true;
    } catch (e) {
      console.warn('[DetailedExtractor] Panel did not open');
      return false;
    }
  },

  /**
   * Close the details panel
   */
  async closeTaskDetails() {
    const closeSelectors = [
      'button[aria-label*="Close"]',
      'button[aria-label="Dismiss"]',
      'button i[data-icon-name="ChromeClose"]',
      'button i[data-icon-name="Cancel"]',
      '[class*="closeButton"]'
    ];

    for (const selector of closeSelectors) {
      const closeBtn = document.querySelector(selector);
      if (closeBtn) {
        const button = closeBtn.closest('button') || closeBtn;
        button.click();
        await DomUtils.sleep(200);
        return true;
      }
    }

    // Try pressing Escape
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await DomUtils.sleep(200);
    return true;
  },

  /**
   * Extract details from the open panel
   */
  extractDetailsFromPanel() {
    const panel = document.querySelector(this.SELECTORS.detailsPanel);
    if (!panel) return null;

    const details = {
      description: '',
      notes: '',
      checklist: [],
      bucket: '',
      labels: [],
      attachments: [],
      comments: []
    };

    // Extract description/notes
    const descSelectors = [
      '[aria-label*="Notes"] textarea',
      '[aria-label*="Description"] textarea',
      '[class*="description"] textarea',
      '[class*="notes"] textarea',
      '[contenteditable="true"]'
    ];

    for (const selector of descSelectors) {
      const el = panel.querySelector(selector);
      if (el) {
        details.description = el.value || el.textContent?.trim() || '';
        break;
      }
    }

    // Extract bucket
    const bucketSelectors = [
      '[aria-label*="Bucket"] input',
      '[aria-label*="Bucket"]',
      '[class*="bucket"] input',
      '[class*="Bucket"]'
    ];

    for (const selector of bucketSelectors) {
      const el = panel.querySelector(selector);
      if (el) {
        details.bucket = el.value || el.textContent?.trim() || '';
        if (details.bucket) break;
      }
    }

    // Extract checklist items
    const checklistContainer = panel.querySelector('[class*="checklist"], [class*="Checklist"]');
    if (checklistContainer) {
      const items = checklistContainer.querySelectorAll('[class*="checklistItem"], [role="listitem"], [class*="item"]');
      items.forEach(item => {
        const checkbox = item.querySelector('input[type="checkbox"]');
        const textEl = item.querySelector('input[type="text"], span, [class*="title"]');
        const text = textEl?.value || textEl?.textContent?.trim();

        if (text) {
          details.checklist.push({
            title: text,
            isChecked: checkbox?.checked || false
          });
        }
      });
    }

    // Try alternative checklist extraction
    if (details.checklist.length === 0) {
      const allCheckboxes = panel.querySelectorAll('input[type="checkbox"]');
      allCheckboxes.forEach(checkbox => {
        const container = checkbox.closest('[class*="item"], [role="listitem"], div');
        if (container) {
          const textEl = container.querySelector('input[type="text"], span:not(:empty)');
          const text = textEl?.value || textEl?.textContent?.trim();
          if (text && text.length < 500) {
            details.checklist.push({
              title: text,
              isChecked: checkbox.checked
            });
          }
        }
      });
    }

    // Extract assigned users
    const personaSelectors = [
      '[class*="persona"]',
      '[class*="Persona"]',
      '[class*="assignee"]',
      '[aria-label*="Assigned"]'
    ];

    const assignedUsers = [];
    for (const selector of personaSelectors) {
      const personas = panel.querySelectorAll(selector);
      personas.forEach(p => {
        const name = p.getAttribute('aria-label') || p.textContent?.trim();
        if (name && name.length < 100 && !assignedUsers.includes(name)) {
          assignedUsers.push(name);
        }
      });
      if (assignedUsers.length > 0) break;
    }
    details.assignedUsers = assignedUsers;

    // Extract labels/tags
    const labelSelectors = ['[class*="label"]', '[class*="tag"]', '[class*="Label"]'];
    for (const selector of labelSelectors) {
      const labels = panel.querySelectorAll(selector);
      labels.forEach(label => {
        const text = label.textContent?.trim();
        if (text && text.length < 50) {
          details.labels.push(text);
        }
      });
    }

    return details;
  },

  /**
   * Extract full details for a single task
   */
  async extractTaskWithDetails(row, rowIndex, onProgress = null) {
    // Get basic info first
    const basicTask = BasicExtractor.extractTaskFromRow(row, rowIndex);
    if (!basicTask) return null;

    // Open details panel
    const panelOpened = await this.openTaskDetails(row);
    if (!panelOpened) {
      basicTask.detailsAvailable = false;
      return basicTask;
    }

    // Extract panel details
    const details = this.extractDetailsFromPanel();

    // Close panel
    await this.closeTaskDetails();

    // Merge basic + detailed info
    return {
      ...basicTask,
      description: details?.description || '',
      notes: details?.notes || '',
      checklist: details?.checklist || [],
      bucket: details?.bucket || basicTask.bucketName,
      labels: details?.labels || [],
      assignedUsers: details?.assignedUsers || basicTask.assignedTo,
      detailsAvailable: true,
      source: 'detailed'
    };
  },

  /**
   * Extract all tasks with full details (slow but complete)
   */
  async extractAllTasksWithDetails(options = {}) {
    const {
      scrollToLoad = true,
      delayBetweenTasks = 500,
      onProgress = null,
      abortSignal = null
    } = options;

    console.log('[DetailedExtractor] Starting detailed extraction...');

    // Scroll to load all tasks first
    if (scrollToLoad) {
      if (onProgress) onProgress({ status: 'scrolling', message: 'Loading all tasks...' });
      await DomUtils.scrollToLoadAll();
    }

    const rows = DomUtils.getTaskRows();
    const tasks = [];
    const errors = [];

    if (onProgress) {
      onProgress({
        status: 'extracting',
        message: `Found ${rows.length} tasks. Opening each task panel...`,
        total: rows.length,
        current: 0
      });
    }

    for (let i = 0; i < rows.length; i++) {
      // Check for abort
      if (abortSignal?.aborted) {
        console.log('[DetailedExtractor] Extraction aborted');
        break;
      }

      const row = rows[i];

      if (onProgress) {
        onProgress({
          status: 'extracting',
          message: `Processing task ${i + 1} of ${rows.length}...`,
          total: rows.length,
          current: i
        });
      }

      try {
        const task = await this.extractTaskWithDetails(row, i, onProgress);
        if (task) {
          tasks.push(task);
        }
      } catch (error) {
        console.error(`[DetailedExtractor] Error extracting task ${i}:`, error);
        errors.push({ index: i, error: error.message });

        // Still try to get basic info
        const basicTask = BasicExtractor.extractTaskFromRow(row, i);
        if (basicTask) {
          basicTask.extractionError = error.message;
          tasks.push(basicTask);
        }
      }

      // Delay between tasks to not overwhelm the UI
      if (i < rows.length - 1) {
        await DomUtils.sleep(delayBetweenTasks);
      }
    }

    console.log(`[DetailedExtractor] Extracted ${tasks.length} tasks with details`);

    if (onProgress) {
      onProgress({
        status: 'complete',
        message: `Extracted ${tasks.length} tasks`,
        total: tasks.length,
        current: tasks.length
      });
    }

    return {
      tasks,
      errors,
      extractionMethod: 'detailed',
      extractedAt: new Date().toISOString()
    };
  }
};

export default DetailedExtractor;
