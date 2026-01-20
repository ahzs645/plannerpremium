/**
 * DOM Utilities for Planner Exporter
 * Common functions for DOM manipulation and extraction
 */

export const DomUtils = {
  /**
   * Sleep for specified milliseconds
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  },

  /**
   * Wait for an element to appear in the DOM
   */
  waitForElement(selector, timeout = 5000) {
    return new Promise((resolve, reject) => {
      const element = document.querySelector(selector);
      if (element) {
        resolve(element);
        return;
      }

      const observer = new MutationObserver((mutations, obs) => {
        const el = document.querySelector(selector);
        if (el) {
          obs.disconnect();
          resolve(el);
        }
      });

      observer.observe(document.body, { childList: true, subtree: true });

      setTimeout(() => {
        observer.disconnect();
        reject(new Error(`Timeout waiting for element: ${selector}`));
      }, timeout);
    });
  },

  /**
   * Wait for element to disappear
   */
  waitForElementToDisappear(selector, timeout = 5000) {
    return new Promise((resolve, reject) => {
      const check = () => {
        if (!document.querySelector(selector)) {
          resolve();
          return true;
        }
        return false;
      };

      if (check()) return;

      const observer = new MutationObserver(() => {
        if (check()) observer.disconnect();
      });

      observer.observe(document.body, { childList: true, subtree: true });

      setTimeout(() => {
        observer.disconnect();
        resolve(); // Resolve anyway after timeout
      }, timeout);
    });
  },

  /**
   * Clean text by removing tooltip/error messages
   */
  cleanText(text) {
    if (!text) return '';
    return text
      .replace(/You can't change.*?(?=\.|$)/gi, '')
      .replace(/summary task/gi, '')
      .trim();
  },

  /**
   * Extract clean text from a cell, filtering out tooltips
   */
  extractCellText(cell) {
    if (!cell) return '';

    // First try input value
    const input = cell.querySelector('input');
    if (input && input.value) {
      return input.value.trim();
    }

    // Walk text nodes, filtering out tooltip text
    let text = '';
    const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT, null, false);
    let node;
    while (node = walker.nextNode()) {
      const nodeText = node.textContent?.trim();
      if (nodeText && !nodeText.includes("You can't change") && !nodeText.includes("summary task")) {
        text += nodeText + ' ';
      }
    }

    return text.trim();
  },

  /**
   * Extract date from text (MM/DD/YYYY format)
   */
  extractDate(text) {
    if (!text) return null;
    const match = text.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
    return match ? match[1] : null;
  },

  /**
   * Extract percentage from text
   */
  extractPercentage(text) {
    if (!text) return 0;
    const match = text.match(/(\d+)\s*%?/);
    return match ? parseInt(match[1], 10) : 0;
  },

  /**
   * Extract duration from text
   */
  extractDuration(text) {
    if (!text) return '';
    const match = text.match(/([\d.]+)\s*(days?|d|hours?|h)/i);
    if (match) {
      const num = parseFloat(match[1]);
      const unit = match[2].toLowerCase().startsWith('h') ? 'hours' : 'days';
      return `${num} ${num === 1 ? unit.slice(0, -1) : unit}`;
    }
    return '';
  },

  /**
   * Map priority text to numeric value
   */
  mapPriority(text) {
    if (!text) return { value: 5, label: 'Medium' };
    const priorityMap = {
      'urgent': { value: 1, label: 'Urgent' },
      'high': { value: 3, label: 'High' },
      'important': { value: 3, label: 'Important' },
      'medium': { value: 5, label: 'Medium' },
      'normal': { value: 5, label: 'Normal' },
      'low': { value: 9, label: 'Low' }
    };
    const match = text.toLowerCase().match(/(urgent|high|important|medium|normal|low)/);
    return match ? priorityMap[match[1]] : { value: 5, label: 'Medium' };
  },

  /**
   * Get all visible task rows
   */
  getTaskRows() {
    const rows = document.querySelectorAll('[role="row"]');
    return Array.from(rows).filter(row => !row.querySelector('[role="columnheader"]'));
  },

  /**
   * Scroll to load all tasks (for virtualized lists)
   */
  async scrollToLoadAll(containerSelector = '[role="grid"]', maxScrolls = 50) {
    const container = document.querySelector(containerSelector);
    if (!container) return;

    let lastRowCount = 0;
    let scrollCount = 0;

    while (scrollCount < maxScrolls) {
      const currentRowCount = document.querySelectorAll('[role="row"]').length;

      if (currentRowCount === lastRowCount) {
        // No new rows loaded, we might be at the end
        break;
      }

      lastRowCount = currentRowCount;
      container.scrollTop = container.scrollHeight;
      await this.sleep(300);
      scrollCount++;
    }

    // Scroll back to top
    container.scrollTop = 0;
    await this.sleep(200);
  }
};

export default DomUtils;
