/**
 * Content Script for Planner Exporter
 * Supports:
 * - Premium Plans (DOM scraping + PSS API)
 * - Basic Plans (Graph API)
 * - Microsoft To Do (Graph API only)
 *
 * Extraction Modes:
 * - Quick: Fast grid-based extraction (basic task info only)
 * - Detailed: Opens each task panel for full info (checklists, notes, etc.)
 */

(function() {
  'use strict';

  // ============================================
  // DOM UTILITIES MODULE
  // ============================================

  const DomUtils = {
    sleep(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    },

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

    extractCellText(cell) {
      if (!cell) return '';

      const input = cell.querySelector('input');
      if (input && input.value) {
        return input.value.trim();
      }

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

    extractDate(text) {
      if (!text) return null;
      const match = text.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
      return match ? match[1] : null;
    },

    extractPercentage(text) {
      if (!text) return 0;
      const match = text.match(/(\d+)\s*%?/);
      return match ? parseInt(match[1], 10) : 0;
    },

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

    getTaskRows() {
      const rows = document.querySelectorAll('[role="row"]');
      return Array.from(rows).filter(row => !row.querySelector('[role="columnheader"]'));
    },

    async scrollToLoadAll(maxScrolls = 50) {
      const container = document.querySelector('[role="grid"]');
      if (!container) return;

      let lastRowCount = 0;
      let scrollCount = 0;

      while (scrollCount < maxScrolls) {
        const currentRowCount = document.querySelectorAll('[role="row"]').length;
        if (currentRowCount === lastRowCount) break;

        lastRowCount = currentRowCount;
        container.scrollTop = container.scrollHeight;
        await this.sleep(300);
        scrollCount++;
      }

      container.scrollTop = 0;
      await this.sleep(200);
    }
  };

  // ============================================
  // BASIC EXTRACTOR MODULE (Quick Mode)
  // ============================================

  const BasicExtractor = {
    extractTaskFromRow(row, rowIndex) {
      const nameCell = row.querySelector('[role="treeitem"]');
      let taskName = '';

      if (nameCell) {
        const input = nameCell.querySelector('input');
        if (input && input.value) {
          taskName = input.value.trim();
        } else {
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

      const expandIcon = nameCell?.querySelector('[data-icon-name*="Chevron"], [class*="chevron"]');
      task.isSummaryTask = !!expandIcon;

      cells.forEach((cell, cellIndex) => {
        const value = DomUtils.extractCellText(cell);

        switch (cellIndex) {
          case 0:
            const checkbox = cell.querySelector('input[type="checkbox"]');
            if (checkbox) {
              task.isComplete = checkbox.checked;
              if (checkbox.checked) task.percentComplete = 100;
            }
            break;
          case 2:
            task.startDateTime = DomUtils.extractDate(value);
            break;
          case 3:
            task.dueDateTime = DomUtils.extractDate(value);
            break;
          case 4:
            if (value && !value.includes("can't change")) {
              task.assignedTo = value.split(',').map(s => s.trim()).filter(s => s && !s.match(/^\+\d+$/));
            }
            break;
          case 5:
            task.duration = DomUtils.extractDuration(value);
            break;
          case 6:
            const pct = DomUtils.extractPercentage(value);
            if (pct > 0) {
              task.percentComplete = pct;
              task.isComplete = pct === 100;
            }
            break;
          case 7:
            const priority = DomUtils.mapPriority(value);
            task.priority = priority.value;
            task.priorityLabel = priority.label;
            break;
        }
      });

      return task;
    },

    async extractAllTasks(options = {}) {
      const { scrollToLoad = false, onProgress = null } = options;

      console.log('[BasicExtractor] Starting quick extraction...');

      if (scrollToLoad) {
        if (onProgress) onProgress({ status: 'scrolling', message: 'Loading all tasks...' });
        await DomUtils.scrollToLoadAll();
      }

      const rows = DomUtils.getTaskRows();
      const tasks = [];

      if (onProgress) {
        onProgress({ status: 'extracting', message: `Found ${rows.length} tasks`, total: rows.length, current: 0 });
      }

      rows.forEach((row, index) => {
        const task = this.extractTaskFromRow(row, index);
        if (task) tasks.push(task);
      });

      console.log(`[BasicExtractor] Extracted ${tasks.length} tasks`);

      if (onProgress) {
        onProgress({ status: 'complete', message: `Done! ${tasks.length} tasks`, total: tasks.length, current: tasks.length });
      }

      return { tasks, extractionMethod: 'basic', extractedAt: new Date().toISOString() };
    }
  };

  // ============================================
  // DETAILED EXTRACTOR MODULE (Full Mode)
  // ============================================

  const DetailedExtractor = {
    async openTaskDetails(row) {
      const nameCell = row.querySelector('[role="treeitem"]');
      if (!nameCell) return false;

      const buttonSelectors = [
        'i[aria-label*="Open details"]',
        'button[aria-label*="Open details"]',
        'i[data-icon-name="Info"]'
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

      button.click();

      try {
        await DomUtils.waitForElement('[role="dialog"], [class*="Panel"][class*="isOpen"], [class*="TaskDetails"]', 3000);
        await DomUtils.sleep(500);
        return true;
      } catch (e) {
        return false;
      }
    },

    async closeTaskDetails() {
      const closeSelectors = [
        'button[aria-label*="Close"]',
        'button[aria-label="Dismiss"]',
        'button i[data-icon-name="ChromeClose"]',
        'button i[data-icon-name="Cancel"]'
      ];

      for (const selector of closeSelectors) {
        const closeBtn = document.querySelector(selector);
        if (closeBtn) {
          const button = closeBtn.closest('button') || closeBtn;
          button.click();
          await DomUtils.sleep(300);
          return true;
        }
      }

      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      await DomUtils.sleep(300);
      return true;
    },

    extractDetailsFromPanel() {
      const panel = document.querySelector('[role="dialog"], [class*="Panel"][class*="isOpen"], [class*="TaskDetails"]');
      if (!panel) return null;

      const details = {
        description: '',
        checklist: [],
        bucket: '',
        labels: [],
        assignedUsers: []
      };

      // Description/Notes
      const descSelectors = ['textarea', '[contenteditable="true"]', '[aria-label*="Notes"]'];
      for (const selector of descSelectors) {
        const el = panel.querySelector(selector);
        if (el) {
          details.description = el.value || el.textContent?.trim() || '';
          if (details.description) break;
        }
      }

      // Bucket
      const bucketSelectors = ['[aria-label*="Bucket"]', '[class*="bucket"]'];
      for (const selector of bucketSelectors) {
        const el = panel.querySelector(selector);
        if (el) {
          details.bucket = el.value || el.textContent?.trim() || '';
          if (details.bucket && !details.bucket.includes('Bucket')) break;
        }
      }

      // Checklist items
      const checklistItems = panel.querySelectorAll('[class*="checklistItem"], [class*="checklist"] [role="listitem"]');
      checklistItems.forEach(item => {
        const checkbox = item.querySelector('input[type="checkbox"]');
        const textEl = item.querySelector('input[type="text"], span, [class*="title"]');
        const text = textEl?.value || textEl?.textContent?.trim();

        if (text && text.length < 500) {
          details.checklist.push({
            title: text,
            isChecked: checkbox?.checked || false
          });
        }
      });

      // Assigned users
      const personas = panel.querySelectorAll('[class*="persona"], [class*="Persona"]');
      personas.forEach(p => {
        const name = p.getAttribute('aria-label') || p.textContent?.trim();
        if (name && name.length < 100) {
          details.assignedUsers.push(name);
        }
      });

      return details;
    },

    async extractTaskWithDetails(row, rowIndex) {
      const basicTask = BasicExtractor.extractTaskFromRow(row, rowIndex);
      if (!basicTask) return null;

      const panelOpened = await this.openTaskDetails(row);
      if (!panelOpened) {
        basicTask.detailsAvailable = false;
        return basicTask;
      }

      const details = this.extractDetailsFromPanel();
      await this.closeTaskDetails();

      return {
        ...basicTask,
        description: details?.description || '',
        checklist: details?.checklist || [],
        bucket: details?.bucket || basicTask.bucketName,
        labels: details?.labels || [],
        assignedUsers: details?.assignedUsers || basicTask.assignedTo,
        detailsAvailable: true,
        source: 'detailed'
      };
    },

    async extractAllTasksWithDetails(options = {}) {
      const { scrollToLoad = true, delayBetweenTasks = 600, onProgress = null } = options;

      console.log('[DetailedExtractor] Starting detailed extraction...');

      if (scrollToLoad) {
        if (onProgress) onProgress({ status: 'scrolling', message: 'Loading all tasks...' });
        await DomUtils.scrollToLoadAll();
      }

      const rows = DomUtils.getTaskRows();
      const tasks = [];

      if (onProgress) {
        onProgress({ status: 'extracting', message: `Found ${rows.length} tasks`, total: rows.length, current: 0 });
      }

      for (let i = 0; i < rows.length; i++) {
        if (onProgress) {
          onProgress({ status: 'extracting', message: `Task ${i + 1} of ${rows.length}`, total: rows.length, current: i });
        }

        try {
          const task = await this.extractTaskWithDetails(rows[i], i);
          if (task) tasks.push(task);
        } catch (error) {
          console.error(`[DetailedExtractor] Error on task ${i}:`, error);
          const basicTask = BasicExtractor.extractTaskFromRow(rows[i], i);
          if (basicTask) {
            basicTask.extractionError = error.message;
            tasks.push(basicTask);
          }
        }

        if (i < rows.length - 1) {
          await DomUtils.sleep(delayBetweenTasks);
        }
      }

      console.log(`[DetailedExtractor] Extracted ${tasks.length} tasks with details`);

      if (onProgress) {
        onProgress({ status: 'complete', message: `Done! ${tasks.length} tasks`, total: tasks.length, current: tasks.length });
      }

      return { tasks, extractionMethod: 'detailed', extractedAt: new Date().toISOString() };
    }
  };

  // ============================================
  // CONTEXT MONITORING
  // ============================================

  let plannerContext = null;

  function injectScript(filename) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = chrome.runtime.getURL(filename);
      script.onload = () => { script.remove(); resolve(); };
      script.onerror = reject;
      (document.head || document.documentElement).appendChild(script);
    });
  }

  function startContextMonitor() {
    const checkBridge = () => {
      const bridge = document.getElementById('planner-exporter-bridge');
      if (bridge) {
        const data = bridge.getAttribute('data-planner-context');
        if (data) {
          try { plannerContext = JSON.parse(data); } catch (e) {}
        }
      }
    };

    // Check immediately and then every second
    checkBridge();
    setInterval(checkBridge, 1000);

    window.addEventListener('message', (event) => {
      if (event.data?.type === 'PLANNER_CONTEXT_UPDATE') {
        plannerContext = event.data.data;

        // Store basic plan context to chrome.storage.local for import page access
        if (plannerContext?.planType === 'basic' && plannerContext?.planId) {
          chrome.storage.local.set({
            plannerBasicPlanId: plannerContext.planId
          }).catch(() => {});
        }
      }

      // Graph API token captured - forward to background.js
      if (event.data?.type === 'PLANNER_TOKEN_CAPTURED') {
        if (plannerContext) plannerContext.token = event.data.token;

        // Store in background.js for centralized access
        chrome.runtime.sendMessage({
          action: 'storeToken',
          type: 'GRAPH',
          token: event.data.token,
          metadata: { timestamp: event.data.timestamp }
        }).catch(() => {});
      }

      // PSS API token captured - forward to background.js
      if (event.data?.type === 'PLANNER_PSS_TOKEN_CAPTURED') {
        if (plannerContext) {
          plannerContext.pssToken = event.data.token;
          plannerContext.pssProjectId = event.data.projectId;
          plannerContext.hasPssAccess = !!(event.data.token && event.data.projectId);
        }

        // Store in background.js for centralized access
        chrome.runtime.sendMessage({
          action: 'storeToken',
          type: 'PSS',
          token: event.data.token,
          metadata: { projectId: event.data.projectId, timestamp: event.data.timestamp }
        }).catch(() => {});
      }

      // To Do API token captured - forward to background.js and store in chrome.storage
      if (event.data?.type === 'TODO_TOKEN_CAPTURED') {
        if (plannerContext) {
          plannerContext.token = event.data.token;
          plannerContext.listId = event.data.listId;
        }

        // Build storage object
        const storageData = {
          todoSubstrateToken: event.data.token,
          todoSubstrateTokenTimestamp: event.data.timestamp || Date.now()
        };

        // Also store X-AnchorMailbox if available (critical for Substrate API)
        if (event.data.anchorMailbox) {
          storageData.todoAnchorMailbox = event.data.anchorMailbox;
          console.log('[PlannerExporter] X-AnchorMailbox captured:', event.data.anchorMailbox);
        }

        // Store directly in chrome.storage for immediate access
        chrome.storage.local.set(storageData).then(() => {
          console.log('[PlannerExporter] To Do Substrate token saved to chrome.storage.local');
        }).catch((err) => {
          console.error('[PlannerExporter] Failed to save token to storage:', err);
        });

        // Also store in background.js for centralized access
        chrome.runtime.sendMessage({
          action: 'storeToken',
          type: 'TODO',
          token: event.data.token,
          metadata: {
            listId: event.data.listId,
            anchorMailbox: event.data.anchorMailbox,
            timestamp: event.data.timestamp
          }
        }).catch(() => {});

        console.log('[PlannerExporter] To Do Substrate token received from page');
      }
    });
  }

  function getCurrentContext() {
    return {
      // Service type
      serviceType: plannerContext?.serviceType || 'planner',

      // Basic info
      token: plannerContext?.token || null,
      planId: plannerContext?.planId || null,
      planType: plannerContext?.planType || null,
      planName: plannerContext?.planName || null,
      hasPlan: !!plannerContext?.planId,
      hasToken: !!plannerContext?.token,

      // PSS API info (Premium Plans)
      pssToken: plannerContext?.pssToken || null,
      pssProjectId: plannerContext?.pssProjectId || null,
      hasPssAccess: plannerContext?.hasPssAccess || false,

      // To Do info
      listId: plannerContext?.listId || null,
      listName: plannerContext?.listName || null,
      hasList: !!plannerContext?.listId
    };
  }

  function getPlanName() {
    if (plannerContext?.planName) return plannerContext.planName;
    const selectors = ['h1', '[role="heading"][aria-level="1"]'];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el?.textContent?.trim()) return el.textContent.trim();
    }
    return 'Unknown Plan';
  }

  // ============================================
  // MAIN EXTRACTION FUNCTION (Hybrid Approach)
  // API calls are delegated to background.js service worker
  // ============================================

  async function extractPlan(mode = 'quick', onProgress = null) {
    const context = getCurrentContext();
    console.log('[PlannerExporter] Extraction context:', context);

    // Microsoft To Do - always use Substrate API
    if (context.serviceType === 'todo' && context.token) {
      console.log('[PlannerExporter] To Do service - calling background.js for Substrate API...');
      console.log('[PlannerExporter] Current list name:', context.listName);

      try {
        if (onProgress) {
          onProgress({ status: 'fetching', message: 'Fetching via Substrate API...' });
        }

        // Delegate to background.js service worker
        // Pass the current list name so we only fetch that list's tasks
        const response = await chrome.runtime.sendMessage({
          action: 'fetchToDoList',
          listId: context.listId,
          listName: context.listName, // Pass the current list name from DOM
          token: context.token
        });

        if (response.success) {
          console.log('[PlannerExporter] To Do Substrate API extraction successful');
          return response.data;
        } else {
          throw new Error(response.error || 'To Do Substrate API failed');
        }

      } catch (apiError) {
        console.error('[PlannerExporter] To Do Substrate API failed:', apiError.message);
        throw apiError; // No DOM fallback for To Do
      }
    }

    // Premium Plan with PSS API access - try API via background.js
    if (context.planType === 'premium' && context.hasPssAccess) {
      console.log('[PlannerExporter] Premium plan with PSS access - calling background.js...');
      console.log('[PlannerExporter] PSS Project ID:', context.pssProjectId);

      try {
        if (onProgress) {
          onProgress({ status: 'fetching', message: 'Fetching via PSS API...' });
        }

        // Delegate to background.js service worker
        const response = await chrome.runtime.sendMessage({
          action: 'fetchPremiumPlan',
          projectId: context.pssProjectId,
          token: context.pssToken
        });

        if (response.success) {
          console.log('[PlannerExporter] PSS API extraction successful');
          return response.data;
        } else {
          throw new Error(response.error || 'PSS API failed');
        }

      } catch (apiError) {
        console.warn('[PlannerExporter] PSS API failed, falling back to DOM:', apiError.message);

        if (onProgress) {
          onProgress({ status: 'fallback', message: 'API failed, using DOM extraction...' });
        }
      }
    }

    // Basic Plan with Graph API access - via background.js
    if (context.planType === 'basic' && context.token) {
      console.log('[PlannerExporter] Basic plan with Graph API access - calling background.js...');

      try {
        if (onProgress) {
          onProgress({ status: 'fetching', message: 'Fetching via Graph API...' });
        }

        // Delegate to background.js service worker
        const response = await chrome.runtime.sendMessage({
          action: 'fetchBasicPlan',
          planId: context.planId,
          token: context.token
        });

        if (response.success) {
          console.log('[PlannerExporter] Graph API extraction successful');
          return response.data;
        } else {
          throw new Error(response.error || 'Graph API failed');
        }

      } catch (apiError) {
        console.warn('[PlannerExporter] Graph API failed, falling back to DOM:', apiError.message);

        if (onProgress) {
          onProgress({ status: 'fallback', message: 'API failed, using DOM extraction...' });
        }
      }
    }

    // Fallback: DOM-based extraction (runs in content script)
    console.log('[PlannerExporter] Using DOM extraction, mode:', mode);

    // Notify background.js that DOM extraction is starting
    chrome.runtime.sendMessage({
      action: 'domExtractionStarted',
      method: mode === 'detailed' ? 'dom-detailed' : 'dom-quick'
    }).catch(() => {});

    try {
      let result;
      if (mode === 'detailed') {
        result = await DetailedExtractor.extractAllTasksWithDetails({
          scrollToLoad: true,
          onProgress
        });
      } else {
        result = await BasicExtractor.extractAllTasks({
          scrollToLoad: true,
          onProgress
        });
      }

      // Notify background.js that DOM extraction completed
      chrome.runtime.sendMessage({
        action: 'domExtractionCompleted',
        taskCount: result.tasks?.length || 0
      }).catch(() => {});

      return result;
    } catch (error) {
      // Notify background.js that DOM extraction failed
      chrome.runtime.sendMessage({
        action: 'domExtractionFailed',
        error: error.message
      }).catch(() => {});

      throw error;
    }
  }

  // ============================================
  // MESSAGE HANDLERS
  // ============================================

  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'refreshContext') {
      // Ask the page context to refresh
      window.postMessage({ type: 'PLANNER_EXPORTER_REFRESH' }, '*');
      sendResponse({ success: true });
      return true;
    }

    if (request.action === 'getContext') {
      // Do a fresh check of the bridge before responding
      const bridge = document.getElementById('planner-exporter-bridge');
      console.log('[PlannerExporter] getContext - bridge element:', bridge ? 'found' : 'NOT FOUND');

      if (bridge) {
        const data = bridge.getAttribute('data-planner-context');
        console.log('[PlannerExporter] getContext - bridge data:', data ? 'has data' : 'NO DATA');
        if (data) {
          try {
            plannerContext = JSON.parse(data);
            console.log('[PlannerExporter] getContext - parsed context:', plannerContext);
          } catch (e) {
            console.error('[PlannerExporter] getContext - parse error:', e);
          }
        }
      }

      const ctx = getCurrentContext();

      // For To Do service, get the Substrate token from chrome.storage
      if (ctx.serviceType === 'todo') {
        chrome.storage.local.get(['todoSubstrateToken', 'todoSubstrateTokenTimestamp'], (data) => {
          if (data.todoSubstrateToken) {
            ctx.token = data.todoSubstrateToken;
            ctx.hasToken = true;
            ctx.tokenSource = 'substrate-storage';
            console.log('[PlannerExporter] getContext - using Substrate token from storage');
          } else {
            console.log('[PlannerExporter] getContext - no Substrate token in storage');
          }
          console.log('[PlannerExporter] getContext - returning:', ctx);
          sendResponse(ctx);
        });
        return true; // Keep channel open for async response
      }

      console.log('[PlannerExporter] getContext - returning:', ctx);
      sendResponse(ctx);
      return true;
    }

    if (request.action === 'extractPlan') {
      const mode = request.mode || 'quick';
      const context = getCurrentContext();

      // Handle To Do service
      if (context.serviceType === 'todo') {
        if (!context.token) {
          sendResponse({ success: false, error: 'No token detected. Please interact with the page first.' });
          return true;
        }
        // For To Do, we don't require a listId - we can fetch all lists
      } else if (!context.planId && context.planType !== 'premium') {
        // Try to detect plan anyway for premium
        if (!window.location.pathname.includes('plan')) {
          sendResponse({ success: false, error: 'No plan detected. Navigate to a plan first.' });
          return true;
        }
      }

      // Send progress updates back to the popup
      const onProgress = (progress) => {
        chrome.runtime.sendMessage({
          action: 'extractionProgress',
          progress: progress
        }).catch(() => {}); // Ignore errors if popup is closed
      };

      extractPlan(mode, onProgress)
        .then(data => {
          const isToDoService = context.serviceType === 'todo';
          const exportData = {
            ...data,
            serviceType: context.serviceType || 'planner',
            plan: data.plan || (isToDoService
              ? { id: context.listId, title: context.listName || 'To Do List' }
              : { id: context.planId, title: getPlanName() }),
            planName: isToDoService ? (context.listName || 'To Do List') : getPlanName(),
            planType: isToDoService ? 'todo' : (context.planType || 'premium'),
            buckets: data.buckets || [],
            bucketMap: data.bucketMap || {},
            detailsMap: data.detailsMap || {},
            exportedAt: new Date().toISOString()
          };

          chrome.storage.local.set({ plannerExportData: exportData }, () => {
            sendResponse({
              success: true,
              taskCount: data.tasks?.length || 0,
              mode: data.extractionMethod || mode,
              data: exportData
            });
          });
        })
        .catch(error => {
          console.error('[PlannerExporter] Export error:', error);
          sendResponse({ success: false, error: error.message });
        });

      return true;
    }

    if (request.action === 'addTask') {
      const context = getCurrentContext();
      if (context.planType === 'premium') {
        sendResponse({ success: false, error: 'Add task not supported for Premium Plans. Use the UI directly.' });
        return true;
      }
      // ... existing add task logic for basic plans
      sendResponse({ success: false, error: 'Add task requires basic plan with API access.' });
      return true;
    }

    if (request.action === 'getBuckets') {
      const buckets = [];
      const headers = document.querySelectorAll('[class*="bucketHeader"], [class*="BucketHeader"]');
      headers.forEach((h, i) => {
        const name = h.textContent?.trim();
        if (name) buckets.push({ id: `bucket-${i}`, name });
      });
      sendResponse({ success: true, buckets });
      return true;
    }

    // ============================================
    // TO DO PAGE API HANDLERS
    // These forward API calls to the page context where CORS doesn't apply
    // ============================================

    if (request.action === 'getToDoListsViaPage') {
      callPageApi('TODO_API_GET_LISTS', {})
        .then(result => sendResponse(result))
        .catch(err => sendResponse({ success: false, error: err.message }));
      return true;
    }

    if (request.action === 'createToDoTaskViaPage') {
      callPageApi('TODO_API_CREATE_TASK', {
        listId: request.listId,
        taskData: request.taskData
      })
        .then(result => sendResponse(result))
        .catch(err => sendResponse({ success: false, error: err.message }));
      return true;
    }

    if (request.action === 'addToDoSubtaskViaPage') {
      callPageApi('TODO_API_ADD_SUBTASK', {
        taskId: request.taskId,
        text: request.text,
        isCompleted: request.isCompleted || false
      })
        .then(result => sendResponse(result))
        .catch(err => sendResponse({ success: false, error: err.message }));
      return true;
    }
  });

  // ============================================
  // PAGE API COMMUNICATION
  // ============================================

  // Store pending page API requests
  const pendingPageApiRequests = new Map();
  let pageApiRequestId = 0;

  // Listen for responses from the page API
  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    if (event.data?.type !== 'TODO_API_RESPONSE') return;

    const { requestId, success, data, error } = event.data;
    const pending = pendingPageApiRequests.get(requestId);

    if (pending) {
      pendingPageApiRequests.delete(requestId);
      if (success) {
        pending.resolve({ success: true, data });
      } else {
        pending.reject(new Error(error || 'Unknown error'));
      }
    }
  });

  // Call the page API and wait for response
  function callPageApi(type, data, timeout = 30000) {
    return new Promise((resolve, reject) => {
      const requestId = `req_${++pageApiRequestId}_${Date.now()}`;

      // Set up timeout
      const timeoutId = setTimeout(() => {
        pendingPageApiRequests.delete(requestId);
        reject(new Error('Page API request timed out'));
      }, timeout);

      // Store the pending request
      pendingPageApiRequests.set(requestId, {
        resolve: (result) => {
          clearTimeout(timeoutId);
          resolve(result);
        },
        reject: (err) => {
          clearTimeout(timeoutId);
          reject(err);
        }
      });

      // Post message to page context
      window.postMessage({
        type: type,
        requestId: requestId,
        ...data
      }, '*');
    });
  }

  // ============================================
  // INITIALIZATION
  // ============================================

  async function init() {
    console.log('[PlannerExporter] Init starting...');
    try {
      await injectScript('fetchOverride.js');
      console.log('[PlannerExporter] fetchOverride.js injected');
      await injectScript('contextBridge.js');
      console.log('[PlannerExporter] contextBridge.js injected');

      // Inject todoPageApi.js for To Do pages
      if (window.location.hostname.includes('to-do.office.com')) {
        await injectScript('todoPageApi.js');
        console.log('[PlannerExporter] todoPageApi.js injected for To Do');
      }

      startContextMonitor();
      console.log('[PlannerExporter] Context monitor started');
    } catch (err) {
      console.error('[PlannerExporter] Init error:', err);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
