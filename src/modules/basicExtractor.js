/**
 * Basic Extractor for Planner Exporter
 * Quick extraction from grid/table view without opening detail panels
 * Best for: Fast exports when you only need basic task info
 */

import { DomUtils } from './domUtils.js';

export const BasicExtractor = {
  /**
   * Extract basic task info from a single row
   */
  extractTaskFromRow(row, rowIndex) {
    // Task name is in [role="treeitem"]
    const nameCell = row.querySelector('[role="treeitem"]');
    let taskName = '';

    if (nameCell) {
      // Try input first
      const input = nameCell.querySelector('input');
      if (input && input.value) {
        taskName = input.value.trim();
      } else {
        // Get text from spans, excluding button text
        const spans = nameCell.querySelectorAll('span');
        for (const span of spans) {
          if (span.closest('button')) continue;
          const text = span.textContent?.trim();
          if (text && !text.includes('Open details') && !text.includes('More options') && text.length > 0) {
            taskName = text;
            break;
          }
        }
      }
    }

    if (!taskName) return null;

    // Get gridcells for other data
    // Index mapping:
    // 0: Checkbox, 1: Spacer, 2: Start, 3: Due, 4: Assigned, 5: Duration, 6: %, 7: Priority
    const cells = row.querySelectorAll('[role="gridcell"]');

    const task = {
      id: `task-${rowIndex}`,
      title: taskName,
      bucketName: '',
      startDateTime: null,
      dueDateTime: null,
      assignedTo: [],
      duration: '',
      percentComplete: 0,
      priority: 5,
      priorityLabel: 'Medium',
      isComplete: false,
      isSummaryTask: false,
      source: 'grid'
    };

    // Detect if this is a summary task (has expand/collapse icon)
    const expandIcon = nameCell?.querySelector('[data-icon-name*="Chevron"], [class*="chevron"]');
    task.isSummaryTask = !!expandIcon;

    cells.forEach((cell, cellIndex) => {
      const value = DomUtils.extractCellText(cell);

      switch (cellIndex) {
        case 0: // Checkbox
          const checkbox = cell.querySelector('input[type="checkbox"]');
          if (checkbox) {
            task.isComplete = checkbox.checked;
            if (checkbox.checked) task.percentComplete = 100;
          }
          break;

        case 2: // Start Date
          task.startDateTime = DomUtils.extractDate(value);
          break;

        case 3: // Due Date
          task.dueDateTime = DomUtils.extractDate(value);
          break;

        case 4: // Assigned To
          if (value && !value.includes("can't change")) {
            task.assignedTo = value
              .split(',')
              .map(s => s.trim())
              .filter(s => s && !s.match(/^\+\d+$/));
          }
          break;

        case 5: // Duration
          task.duration = DomUtils.extractDuration(value);
          break;

        case 6: // % Complete
          const pct = DomUtils.extractPercentage(value);
          if (pct > 0) {
            task.percentComplete = pct;
            task.isComplete = pct === 100;
          }
          break;

        case 7: // Priority
          const priority = DomUtils.mapPriority(value);
          task.priority = priority.value;
          task.priorityLabel = priority.label;
          break;
      }
    });

    return task;
  },

  /**
   * Extract all tasks from the grid (quick mode)
   */
  async extractAllTasks(options = {}) {
    const {
      scrollToLoad = false,
      onProgress = null
    } = options;

    console.log('[BasicExtractor] Starting quick extraction...');

    // Optionally scroll to load all tasks
    if (scrollToLoad) {
      if (onProgress) onProgress({ status: 'scrolling', message: 'Loading all tasks...' });
      await DomUtils.scrollToLoadAll();
    }

    const rows = DomUtils.getTaskRows();
    const tasks = [];

    if (onProgress) onProgress({ status: 'extracting', total: rows.length, current: 0 });

    rows.forEach((row, index) => {
      const task = this.extractTaskFromRow(row, index);
      if (task) {
        tasks.push(task);
      }

      if (onProgress && index % 10 === 0) {
        onProgress({ status: 'extracting', total: rows.length, current: index });
      }
    });

    console.log(`[BasicExtractor] Extracted ${tasks.length} tasks`);

    if (onProgress) onProgress({ status: 'complete', total: tasks.length, current: tasks.length });

    return {
      tasks,
      extractionMethod: 'basic',
      extractedAt: new Date().toISOString()
    };
  },

  /**
   * Extract buckets from DOM (board view)
   */
  extractBuckets() {
    const buckets = [];
    const bucketHeaders = document.querySelectorAll(
      '[class*="bucketHeader"], [class*="BucketHeader"], [class*="columnHeader"]'
    );

    bucketHeaders.forEach((header, index) => {
      const name = header.textContent?.trim();
      if (name) {
        buckets.push({ id: `bucket-${index}`, name });
      }
    });

    return buckets;
  },

  /**
   * Get plan name from DOM
   */
  getPlanName() {
    const selectors = [
      'h1',
      '[role="heading"][aria-level="1"]',
      '[class*="planName"]',
      '[class*="PlanName"]'
    ];

    for (const selector of selectors) {
      const el = document.querySelector(selector);
      if (el && el.textContent?.trim()) {
        return el.textContent.trim();
      }
    }

    // Try document title
    const title = document.title;
    if (title) {
      return title.replace(/\s*[-|]\s*Microsoft Planner.*$/i, '').trim();
    }

    return 'Unknown Plan';
  }
};

export default BasicExtractor;
