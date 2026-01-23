/**
 * Context Bridge for Planner Exporter
 * Runs in page context to extract plan info and tokens
 * Supports:
 * - Premium Plans (Project for the Web)
 * - Basic Plans (standard Planner)
 * - Microsoft To Do (to-do.office.com)
 */

(function() {
  'use strict';

  // Create bridge element to expose data to content script
  function createBridgeElement() {
    let bridge = document.getElementById('planner-exporter-bridge');
    if (!bridge) {
      bridge = document.createElement('div');
      bridge.id = 'planner-exporter-bridge';
      bridge.style.display = 'none';
      document.documentElement.appendChild(bridge);
    }
    return bridge;
  }

  // Detect service type from URL
  function detectServiceType() {
    const hostname = window.location.hostname;
    if (hostname.includes('to-do.office.com')) return 'todo';
    if (hostname.includes('tasks.office.com') || hostname.includes('planner.cloud.microsoft')) return 'planner';
    return null;
  }

  // Extract To Do context from URL
  // URL patterns: /tasks/id/{listId}, /tasks/{listId}, /tasks
  function extractToDoContext() {
    if (detectServiceType() !== 'todo') return null;

    const pathname = window.location.pathname;
    let listId = null;

    // Try various URL patterns
    // Pattern: /tasks/id/{listId}
    let match = pathname.match(/\/tasks\/id\/([^\/\?]+)/);
    if (match) {
      listId = match[1];
    }

    // Pattern: /tasks/{listId} (where listId is not 'id', 'inbox', 'myday', etc.)
    if (!listId) {
      match = pathname.match(/\/tasks\/([^\/\?]+)/);
      if (match && !['id', 'inbox', 'myday', 'planned', 'all', 'completed', 'assigned'].includes(match[1].toLowerCase())) {
        listId = match[1];
      }
    }

    // Also check for captured list ID from API calls
    if (!listId && window.__todoListId) {
      listId = window.__todoListId;
    }

    return {
      serviceType: 'todo',
      listId: listId,
      listName: getToDoListNameFromDom()
    };
  }

  // Get To Do list name from DOM
  function getToDoListNameFromDom() {
    // Try various selectors for list name in To Do
    // The h1 heading typically shows the current list name
    const selectors = [
      'h1',
      '[role="heading"][aria-level="1"]',
      '[data-automationid="ListHeaderName"]',
      '[class*="listName"]',
      '[class*="ListName"]',
      '[class*="ListHeader"] span',
      '[class*="listHeader"] span'
    ];

    for (const selector of selectors) {
      try {
        const els = document.querySelectorAll(selector);
        for (const el of els) {
          const text = el.textContent?.trim();
          // Filter out generic titles and ensure it's a real list name
          if (text && text.length > 0 && text.length < 100 &&
              !text.toLowerCase().includes('microsoft to do') &&
              !text.toLowerCase().includes('to-do')) {
            console.log('[Planner Exporter] Found To Do list name:', text);
            return text;
          }
        }
      } catch (e) {}
    }

    return null;
  }

  // Extract plan context from URL
  // Supports: /webui/premiumplan/{id}/... and /webui/basicplan/{id}/...
  function extractPlanContext() {
    // Check if we're on To Do instead
    if (detectServiceType() === 'todo') {
      return extractToDoContext();
    }

    const pathname = window.location.pathname;
    const segments = pathname.split('/').filter(Boolean);

    let planType = null;
    let planId = null;
    let orgId = null;

    // Check for premiumplan (Project for the Web)
    const premiumIndex = segments.indexOf('premiumplan');
    if (premiumIndex !== -1 && segments[premiumIndex + 1]) {
      planType = 'premium';
      planId = segments[premiumIndex + 1].split('?')[0]; // Remove any query params
    }

    // Check for basicplan (standard Planner)
    const basicIndex = segments.indexOf('basicplan');
    if (basicIndex !== -1 && segments[basicIndex + 1]) {
      planType = 'basic';
      planId = segments[basicIndex + 1].split('?')[0];
    }

    // Also check for just "plan" in URL (older format)
    if (!planId) {
      const planIndex = segments.indexOf('plan');
      if (planIndex !== -1 && segments[planIndex + 1]) {
        planType = 'basic';
        planId = segments[planIndex + 1].split('?')[0];
      }
    }

    // Extract org ID
    const orgIndex = segments.indexOf('org');
    if (orgIndex !== -1 && segments[orgIndex + 1]) {
      orgId = segments[orgIndex + 1].split('?')[0];
    }

    // Extract tenant ID from query params
    const urlParams = new URLSearchParams(window.location.search);
    const tenantId = urlParams.get('tid');

    return { planType, planId, orgId, tenantId };
  }

  // Get Graph API access token - prioritize captured token from fetch intercept
  function getAccessToken() {
    // Method 1: Check for captured global token (most reliable for encrypted MSAL)
    if (window.__plannerGraphToken) {
      return {
        token: window.__plannerGraphToken,
        capturedAt: window.__plannerGraphTokenCapturedAt,
        source: 'fetch-intercept'
      };
    }

    // Method 2: Try localStorage (works for older MSAL format)
    for (const key of Object.keys(localStorage)) {
      if (key.toLowerCase().includes('accesstoken')) {
        try {
          const data = JSON.parse(localStorage.getItem(key));

          // Check for graph.microsoft.com scope
          const target = (data.target || '').toLowerCase();
          const keyLower = key.toLowerCase();

          if ((target.includes('graph.microsoft.com') || keyLower.includes('graph.microsoft.com'))) {
            // Old format: plaintext secret
            if (data.secret && typeof data.secret === 'string' && data.secret.length > 100) {
              // Check expiration
              if (data.expiresOn) {
                const expiresOn = parseInt(data.expiresOn, 10);
                if (expiresOn * 1000 < Date.now()) {
                  continue; // Skip expired
                }
              }
              return {
                token: data.secret,
                expiresOn: data.expiresOn,
                source: 'localStorage'
              };
            }
          }
        } catch (e) {
          // Not JSON or malformed
        }
      }
    }

    return null;
  }

  // Get PSS API token and project ID (for Premium Plans)
  function getPssContext() {
    if (window.__plannerPssToken) {
      return {
        token: window.__plannerPssToken,
        capturedAt: window.__plannerPssTokenCapturedAt,
        projectId: window.__plannerPssProjectId,
        source: 'fetch-intercept'
      };
    }
    return null;
  }

  // Get To Do token context
  function getToDoContext() {
    if (window.__todoToken) {
      return {
        token: window.__todoToken,
        capturedAt: window.__todoTokenCapturedAt,
        listId: window.__todoListId,
        source: 'fetch-intercept'
      };
    }
    // Also check if Graph token exists (To Do uses same token)
    if (window.__plannerGraphToken) {
      return {
        token: window.__plannerGraphToken,
        capturedAt: window.__plannerGraphTokenCapturedAt,
        listId: window.__todoListId,
        source: 'graph-token'
      };
    }
    return null;
  }

  // Build PSS project identifier from context
  // Format: msxrm_{dynamicsOrg}_{planId}
  function buildPssProjectId(planId, orgId) {
    if (!planId) return null;

    // If we already captured a project ID from API calls, use that
    if (window.__plannerPssProjectId) {
      return window.__plannerPssProjectId;
    }

    // If we captured the Dynamics org from API URLs, build the project ID
    if (window.__plannerDynamicsOrg) {
      const projectId = `msxrm_${window.__plannerDynamicsOrg}_${planId}`;
      console.log('[Planner Exporter] Built project ID from captured org:', projectId);
      return projectId;
    }

    // Try to find Dynamics org from page - look for crm.dynamics.com patterns
    // Check script tags, data attributes, or window objects
    const scripts = document.querySelectorAll('script');
    for (const script of scripts) {
      const content = script.textContent || '';
      const orgMatch = content.match(/(org[a-f0-9]+\.crm\d*\.dynamics\.com)/i);
      if (orgMatch) {
        const dynamicsOrg = orgMatch[1];
        const projectId = `msxrm_${dynamicsOrg}_${planId}`;
        console.log('[Planner Exporter] Built project ID from page:', projectId);
        return projectId;
      }
    }

    return null;
  }

  // Get plan name from DOM
  function getPlanNameFromDom() {
    // Try various selectors for plan name
    const selectors = [
      '[data-automation-id="planTitle"]',
      '[class*="planName"]',
      '[class*="PlanName"]',
      '.plan-title',
      '[aria-label*="plan name"]',
      '[role="heading"][aria-level="1"]',
      'h1'
    ];

    for (const selector of selectors) {
      try {
        const els = document.querySelectorAll(selector);
        for (const el of els) {
          const text = el.textContent?.trim();
          if (text && text.length > 0 && text.length < 100 && !text.includes('Planner')) {
            return text;
          }
        }
      } catch (e) {
        // Selector might be invalid
      }
    }

    // Fallback: try page title
    const title = document.title;
    if (title) {
      // Remove common suffixes
      const cleaned = title
        .replace(/\s*[-|]\s*Microsoft Planner.*$/i, '')
        .replace(/\s*[-|]\s*Planner.*$/i, '')
        .trim();
      if (cleaned && cleaned.length < 100) {
        return cleaned;
      }
    }

    return null;
  }

  // Update bridge with current data
  function updateBridge() {
    const bridge = createBridgeElement();

    const serviceType = detectServiceType();
    const tokenInfo = getAccessToken();
    const pssContext = getPssContext();
    const todoContext = getToDoContext();

    let data;

    // Handle To Do service
    if (serviceType === 'todo') {
      const todoInfo = extractToDoContext();
      const listName = todoInfo?.listName || getToDoListNameFromDom();

      data = {
        // Service type
        serviceType: 'todo',

        // To Do list info
        listId: todoInfo?.listId || todoContext?.listId || null,
        listName: listName,

        // For compatibility with existing code
        planId: null,
        planType: null,
        planName: listName,

        // Token info (To Do uses Graph API)
        token: todoContext?.token || tokenInfo?.token || null,
        tokenSource: todoContext?.source || tokenInfo?.source || null,
        tokenCapturedAt: todoContext?.capturedAt || tokenInfo?.capturedAt || null,

        // No PSS for To Do
        pssToken: null,
        pssProjectId: null,
        hasPssAccess: false,

        // Meta
        url: window.location.href,
        timestamp: Date.now()
      };

      console.log('[Planner Exporter] To Do context:', data);
    } else {
      // Handle Planner service (existing logic)
      const planContext = extractPlanContext();
      const planName = getPlanNameFromDom();

      // Try to get or build the PSS project ID
      let pssProjectId = pssContext?.projectId || null;
      if (!pssProjectId && planContext?.planType === 'premium' && planContext?.planId) {
        pssProjectId = buildPssProjectId(planContext.planId, planContext.orgId);
      }

      data = {
        // Service type
        serviceType: 'planner',

        // Plan info
        planId: planContext?.planId || null,
        planType: planContext?.planType || null, // 'premium' or 'basic'
        orgId: planContext?.orgId || null,
        tenantId: planContext?.tenantId || null,
        planName: planName,

        // To Do fields (null for Planner)
        listId: null,
        listName: null,

        // Graph API token info
        token: tokenInfo?.token || null,
        tokenSource: tokenInfo?.source || null,
        tokenCapturedAt: tokenInfo?.capturedAt || null,

        // PSS API info (Premium Plans)
        pssToken: pssContext?.token || null,
        pssProjectId: pssProjectId,
        pssTokenCapturedAt: pssContext?.capturedAt || null,
        hasPssAccess: !!(pssContext?.token && pssProjectId),

        // Meta
        url: window.location.href,
        timestamp: Date.now()
      };

      console.log('[Planner Exporter] Context bridge data:', data);
      console.log('[Planner Exporter] PSS fields - projectId:', pssProjectId, 'hasPssAccess:', !!(pssContext?.token && pssProjectId));
    }

    bridge.setAttribute('data-planner-context', JSON.stringify(data));

    // Dispatch event for content script
    bridge.dispatchEvent(new CustomEvent('planner-context-updated', { detail: data }));

    // Also post message for cross-context communication
    window.postMessage({ type: 'PLANNER_CONTEXT_UPDATE', data: data }, '*');
  }

  // Initial update immediately and then with a short delay
  updateBridge();
  setTimeout(updateBridge, 500);

  // Update periodically
  setInterval(updateBridge, 2000);

  // Listen for URL changes (SPA navigation)
  let lastUrl = window.location.href;
  const observer = new MutationObserver(() => {
    if (window.location.href !== lastUrl) {
      lastUrl = window.location.href;
      setTimeout(updateBridge, 500);
    }
  });

  if (document.body) {
    observer.observe(document.body, { childList: true, subtree: true });
  } else {
    document.addEventListener('DOMContentLoaded', () => {
      observer.observe(document.body, { childList: true, subtree: true });
    });
  }

  // Listen for manual refresh requests
  window.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'PLANNER_EXPORTER_REFRESH') {
      updateBridge();
    }
  });

  console.log('[Planner Exporter] Context bridge initialized');
})();
